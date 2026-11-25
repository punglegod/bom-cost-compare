# bom-cost-compare（BOM 价差对比与逐料小结）

给 AI agent 使用的技能：核对 Excel 成本明细中现款与参考款、或同一产品前后两版 BOM 的价差。按成品编码定位参考明细，逐料核对编码、规格、用量和核算单价，给每种差异物料写带双方料号和金额的一句话小结，并核平整机含税差额。适用于采购建价复核、改版对比和建价小结。

## 结构

- `SKILL.md` — 技能正文（找两端、配对归因、核平含税价差、交付格式）
- `scripts/compare_bom.py` — 只读核对脚本，抽取「汇总＋子项明细」版式的两端数据（需 Python 3.8+ 与 openpyxl）
- `references/material-summaries.md` — 逐料小结句式示例
- `agents/openai.yaml` — Codex 等工具的技能声明

## 用法

把整个技能文件夹、主表和参考表交给 agent，并指定现成品编码；支持技能发现的工具可直接调用 `$bom-cost-compare`。命令示例与口径见 `SKILL.md`。
