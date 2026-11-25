#!/usr/bin/env python3
"""Read cached BOM cost workbooks and emit a current-minus-reference JSON audit.

Requires Python 3.8+ and openpyxl. Exit 0 for a reconciled comparison, 1 when
the audit finds missing data or a discrepancy, and 2 for an input/layout error.
"""

import argparse
from collections import Counter, defaultdict
from contextlib import ExitStack, closing
from datetime import date, time, timedelta
from decimal import Decimal, InvalidOperation
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import unicodedata
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


# Only absorbs spreadsheet floating-point noise in reconciliation checks.
# All emitted residuals and all line changes retain the source precision.
TOLERANCE = Decimal("0.00000001")
ZERO = Decimal(0)
SUMMARY = {
    "code": ("编码", "产品编码", "产品代码"), "name": ("名称", "产品名称"),
    "spec": ("规格型号",), "date": ("报价日期",),
    "material": ("物料费用", "物料费用(元)", "材料费用", "材料费用(元)", "材料成本", "材料成本(元)"),
    "labor": ("工费", "工费(填写)", "工费(元)"),
    "loss": ("损耗",), "markup": ("加成",), "pretax": ("不含税成本", "未税成本"),
    "tax_inclusive": ("含税成本",), "reference_code": ("参考编码", "参考SKU"),
    "reference_price": ("参考价格", "参考SKU价格"),
}
DETAIL = {
    "product": ("产品代码", "产品编码"), "code": ("子项物料编码",),
    "level": ("BOM层级",), "name": ("物料名称",), "spec": ("规格型号",),
    "attribute": ("物料属性",), "unit": ("单位",), "type": ("子项类型",),
    "quantity": ("核对数量",), "numerator": ("用量:分子",), "denominator": ("用量:分母",),
    "bom_unit_price": ("单价",), "price_list_unit_price": ("价目表",),
    "cost": ("核算成本",), "note": ("备注",),
}
CONVERSION = {
    "key": ("行匹配键",), "product": ("产品编码", "产品代码"), "code": ("子项物料编码",),
    "divisor": ("核算除数",), "pricing_unit": ("计价单位",),
}
TOTALS = ("material", "labor", "loss", "markup", "pretax", "tax_inclusive")


class AuditError(ValueError):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


def header_name(value):
    value = unicodedata.normalize("NFKC", str(value or ""))
    value = re.sub(r"\s+", "", value).casefold()
    # Recognize the amount columns' actual rate annotations, not arbitrary
    # qualifiers such as "(万元)" or a percentage-only input field.
    return re.sub(r"^(损耗|加成|含税成本)\((?:默认)?\d+(?:\.\d+)?%\)$", r"\1", value)


def table(sheet, schema, required):
    """Accept leading titles, reordered columns, and explicit header aliases only."""
    aliases = {header_name(name): key for key, names in schema.items() for name in names}
    rows = iter(sheet.values)
    for row_no, row in enumerate(rows, 1):
        matches = defaultdict(list)
        for index, value in enumerate(row):
            key = aliases.get(header_name(value))
            if key:
                matches[key].append(index)
        if required.issubset(matches):
            duplicates = [key for key, indices in matches.items() if len(indices) != 1]
            if duplicates:
                raise AuditError("unsupported_layout", f"{sheet.title}!{row_no} 表头重复 {duplicates}")
            columns = {key: indices[0] for key, indices in matches.items()}
            return row_no, columns, enumerate(rows, row_no + 1)
        if row_no >= 50:
            break
    raise AuditError("unsupported_layout", f"{sheet.title} 前50行未找到所需表头 {sorted(required)}")


def record(row, columns):
    return {key: row[index] if index < len(row) else None for key, index in columns.items()}


def code_text(value):
    # Do not upper-case, strip leading zeroes or remove suffixes.
    if isinstance(value, bool) or value is None:
        return None
    result = str(value).strip()
    if not result or result.startswith("#"):
        return None
    return result


def number(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        value = Decimal(str(value).strip())
    except InvalidOperation:
        return None
    return value if value.is_finite() else None


def subtract(left, right):
    return left - right if left is not None and right is not None else None


def strict_sum(values):
    values = list(values)
    return sum(values, ZERO) if all(v is not None for v in values) else None


def address(sheet, col_index, row):
    return f"{sheet}!{get_column_letter(col_index + 1)}{row}"


def numeric_field(data, key, columns, sheet, row, issues, required):
    raw = data.get(key)
    value = number(raw)
    if value is None and (required or raw not in (None, "")):
        issues.append({"cell": address(sheet, columns[key], row), "field": key,
                       "problem": "缺少缓存数值或值无效", "value": raw})
    return value


def summary_info(book, formulas, product, issues):
    name = "汇总"
    header, cols, rows = table(book[name], SUMMARY, {"code", *TOTALS})
    matches = [(n, record(row, cols)) for n, row in rows if code_text(row[cols["code"]]) == product]
    if len(matches) != 1:
        raise AuditError("selection_error", f"{name} 中 {product} 命中 {len(matches)} 行，需要唯一编码")
    row_no, data = matches[0]
    formula_row = next(formulas[name].iter_rows(min_row=row_no, max_row=row_no, values_only=True))
    result = {"header_row": header, "row": row_no, "values": {}, "cells": {}, "formulas": {}}
    for key, index in cols.items():
        result["cells"][key] = address(name, index, row_no)
        cell = formula_row[index]
        if isinstance(cell, str) and cell.startswith("="):
            result["formulas"][key] = cell
        if key in TOTALS or key == "reference_price":
            value = numeric_field(data, key, cols, name, row_no, issues,
                                  key in TOTALS or bool(code_text(data.get("reference_code"))))
        else:
            value = data[key]
        result["values"][key] = value
    return result


def conversions(book, product, issues):
    name = "规格换算"
    if name not in book.sheetnames:
        return None
    _, cols, rows = table(book[name], CONVERSION, set(CONVERSION))
    result = {}
    for row_no, row in rows:
        data = record(row, cols)
        if code_text(data["product"]) != product:
            continue
        key = code_text(data["key"])
        if key in result:
            issues.append({"cell": address(name, cols["key"], row_no), "problem": "重复行匹配键", "value": key})
            result[key] = None
            continue
        result[key] = {"row": row_no, "code": code_text(data["code"]),
                       "divisor": numeric_field(data, "divisor", cols, name, row_no, issues, True),
                       "pricing_unit": data["pricing_unit"]}
    return result


def detail_info(book, formulas, product, issues):
    name = "子项明细"
    required = {"product", "code", "level", "name", "spec", "attribute", "unit", "quantity",
                "bom_unit_price", "price_list_unit_price", "cost"}
    header, cols, rows = table(book[name], DETAIL, required)
    conversion_map = conversions(book, product, issues)
    formula_rows = formulas[name].iter_rows(min_row=header + 1, values_only=True)
    seen = Counter()
    groups = defaultdict(list)
    roots, unkeyed = [], []
    costs = []
    for (row_no, row), formula_row in zip(rows, formula_rows):
        data = record(row, cols)
        owner = code_text(data["product"])
        if owner != product:
            if owner is None and any(data[k] not in (None, "") for k in ("code", "name", "cost")):
                issues.append({"row": f"{name}!{row_no}", "problem": "缺少产品代码，无法确定所属成品"})
            continue
        material = code_text(data["code"])
        line = {"row": row_no, "code": material}
        expressions = {}
        for field in ("cost", "quantity", "price_list_unit_price"):
            expression = formula_row[cols[field]]
            if isinstance(expression, str) and expression.startswith("="):
                expressions[field] = expression
        line["formula_fields"] = list(expressions)
        for key in ("level", "name", "spec", "attribute", "unit", "type", "note"):
            line[key] = data.get(key)
        assembly = data["attribute"] == "自制" or str(data.get("note")).strip().casefold() == "二级bom"
        if not assembly and not str(data["unit"] or "").strip():
            issues.append({"cell": address(name, cols["unit"], row_no), "problem": "子项缺少单位"})
        for key in ("quantity", "bom_unit_price", "price_list_unit_price", "cost", "numerator", "denominator"):
            if key in cols:
                mandatory = key == "cost" or (key in ("quantity", "price_list_unit_price") and not assembly)
                line[key] = numeric_field(data, key, cols, name, row_no, issues, mandatory)
        cost, qty = line["cost"], line["quantity"]
        costs.append(cost)
        line["effective_cost_per_detail_unit"] = cost / qty if cost is not None and qty else None
        if material is None:
            issues.append({"cell": address(name, cols["code"], row_no), "problem": "子项缺少有效料号"})
            unkeyed.append(line)
            continue
        seen[material] += 1
        if conversion_map is not None:
            match_key = f"{product}|{material}|{seen[material]}"
            conversion = conversion_map.get(match_key)
            line["conversion"] = conversion
            if conversion is None or conversion["code"] != material:
                issues.append({"row": f"{name}!{row_no}", "problem": "规格换算键未唯一匹配", "key": match_key})
            elif not assembly:
                divisor, price = conversion["divisor"], line["price_list_unit_price"]
                if divisor is not None and divisor <= 0:
                    issues.append({"row": f"规格换算!{conversion['row']}", "problem": "核算除数非正"})
                elif divisor and qty is not None and price is not None and cost is not None:
                    residual = cost - qty * price / divisor
                    line["cost_minus_quantity_times_price_over_divisor"] = residual
                    if abs(residual) > TOLERANCE:
                        issues.append({"row": f"{name}!{row_no}", "material": material,
                                       "problem": "核算成本与数量、价目表、除数不符",
                                       "cost_source": "公式缓存" if "cost" in expressions else "单元格常数",
                                       "formulas": expressions, "residual": residual})
        if material == product and str(data["level"]).strip() == "0":
            roots.append(line)
            if cost != ZERO and cost is not None:
                issues.append({"row": f"{name}!{row_no}", "problem": "成品根行有非零核算成本", "value": cost})
        else:
            groups[material].append(line)
    if not costs:
        issues.append({"sheet": name, "problem": "未找到指定成品明细", "product": product})
    return {
        "header_row": header, "columns": {k: get_column_letter(c + 1) for k, c in cols.items()},
        "selected_rows": len(costs), "known_cost_sum": sum((n for n in costs if n is not None), ZERO),
        "complete_cost_sum": strict_sum(costs) if costs else None,
        "root_rows": roots, "unkeyed_rows": unkeyed,
        "materials": {code: aggregate(lines) for code, lines in groups.items()},
    }


def aggregate(lines):
    quantities = defaultdict(list)
    for line in lines:
        quantities[str(line["unit"] or "")].append(line["quantity"])
    return {"cost": strict_sum(line["cost"] for line in lines),
            "quantity_by_unit": {unit: strict_sum(values) for unit, values in quantities.items()},
            "rows": lines}


def read_product(path, product):
    issues = []
    try:
        with ExitStack() as stack:
            book, formulas = [stack.enter_context(closing(load_workbook(
                path, read_only=True, data_only=cached, keep_links=False))) for cached in (True, False)]
            for sheet_name in ("汇总", "子项明细"):
                if sheet_name not in book.sheetnames:
                    raise AuditError("unsupported_layout", f"缺少 {sheet_name} 工作表，现有 {book.sheetnames}")
            summary = summary_info(book, formulas, product, issues)
            detail = detail_info(book, formulas, product, issues)
    except AuditError as exc:
        raise AuditError(exc.status, f"{path} — {exc}") from exc
    except (OSError, BadZipFile, InvalidFileException, ParseError, KeyError, ValueError) as exc:
        raise AuditError("input_error", f"无法读取 {path} — {exc}") from exc
    values = summary["values"]
    checks = {
        "detail_minus_summary_material": subtract(detail["complete_cost_sum"], values["material"]),
        "pretax_minus_components": subtract(values["pretax"], strict_sum(values[k] for k in TOTALS[:4])),
    }
    for key, residual in checks.items():
        if residual is not None and abs(residual) > TOLERANCE:
            issues.append({"check": key, "residual": residual})
    return {"file": str(path.absolute()), "product": product, "summary": summary, "detail": detail,
            "checks": checks, "issues": issues}


def field_values(group, key):
    return Counter(row.get(key) for row in group["rows"])


def compare(current, reference):
    current_values, reference_values = [side["summary"]["values"] for side in (current, reference)]
    deltas = {k: subtract(current_values[k], reference_values[k]) for k in TOTALS}
    reference_total = reference_values["tax_inclusive"]
    actual_delta = deltas["tax_inclusive"]
    current_items, reference_items = [side["detail"]["materials"] for side in (current, reference)]
    changed, current_only, reference_only = [], [], []
    contributions = []
    for code in sorted(current_items.keys() | reference_items.keys()):
        new, old = current_items.get(code), reference_items.get(code)
        delta = subtract(new["cost"] if new else ZERO, old["cost"] if old else ZERO)
        contributions.append(delta)
        item = {"code": code, "cost_delta": delta}
        if new and old:
            fields = [field for field in ("quantity", "name", "spec", "unit", "attribute", "cost",
                      "price_list_unit_price", "bom_unit_price", "level", "numerator", "denominator")
                      if field_values(new, field) != field_values(old, field)]
            for key in ("pricing_unit", "divisor"):
                if Counter((r.get("conversion") or {}).get(key) for r in new["rows"]) != Counter(
                        (r.get("conversion") or {}).get(key) for r in old["rows"]):
                    fields.append(key)
            if fields:
                item["changed_fields"] = fields
                changed.append(item)
        elif new:
            current_only.append(item)
        else:
            reference_only.append(item)
    material_contribution = strict_sum(contributions)
    residual = subtract(deltas["material"], material_contribution)
    needs_review = bool(current["issues"] or reference["issues"] or residual is None or abs(residual) > TOLERANCE)
    sheet_reference = current_values.get("reference_price")
    return {
        "status": "review_required" if needs_review else "reconciled",
        "direction": "current - reference", "absolute_reconciliation_tolerance": TOLERANCE,
        "number_format": "十进制定点字符串，未知为 null", "current": current, "reference": reference,
        "delta": deltas, "tax_inclusive_delta_divided_by_reference":
            actual_delta / reference_total if reference_total and actual_delta is not None else None,
        "reference_price_check": {
            "current_sheet_reference_code": current_values.get("reference_code"),
            "selected_reference_code": reference["product"],
            "current_sheet_reference_price": sheet_reference,
            "selected_reference_tax_inclusive": reference_total,
            "current_tax_inclusive_minus_sheet_reference": subtract(current_values["tax_inclusive"], sheet_reference),
            "sheet_reference_minus_selected_reference": subtract(sheet_reference, reference_total),
        },
        "same_code_changes": changed, "current_only": current_only, "reference_only": reference_only,
        "component_delta_sum": material_contribution,
        "summary_material_delta_minus_component_delta_sum": residual,
    }


def serialize(value):
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    raise TypeError(f"不支持的 JSON 值 {type(value).__name__}")


def emit(result, output):
    text = json.dumps(result, ensure_ascii=False, indent=2, default=serialize) + "\n"
    if output is None:
        sys.stdout.buffer.write(text.encode("utf-8"))
        return
    with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=output.parent,
                                     prefix=".bom-", suffix=".tmp", delete=False) as out:
        temp = Path(out.name)
        try:
            out.write(text)
        except BaseException:
            out.close()
            temp.unlink()
            raise
    try:
        os.replace(temp, output)
    finally:
        if temp.exists():
            temp.unlink()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--current", required=True, type=Path)
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--current-code", required=True)
    parser.add_argument("--reference-code", required=True)
    parser.add_argument("--output", type=Path, help="可选 UTF-8 .json；默认写入标准输出")
    args = parser.parse_args()
    if args.output:
        for src in (args.current, args.reference):
            if args.output.resolve() == src.resolve() or (
                    args.output.exists() and src.exists() and args.output.samefile(src)):
                parser.error("输出路径与输入工作簿相同")
        if args.output.suffix.casefold() != ".json":
            parser.error("--output 文件名需以 .json 结尾")
    try:
        current = read_product(args.current, args.current_code)
        reference = read_product(args.reference, args.reference_code)
        result = compare(current, reference)
        exit_code = 0 if result["status"] == "reconciled" else 1
    except AuditError as exc:
        result, exit_code = {"status": exc.status, "message": str(exc)}, 2
    try:
        emit(result, args.output)
    except OSError as exc:
        parser.exit(2, f"无法写入结果 — {exc}\n")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
