# AI 记账 · 月度账单整理与消费报告（个人/家庭）

*A Chinese personal-finance skill: dedupe multi-source bills (WeChat Pay / Alipay / bank & credit-card) → categorize → monthly spending report.*

把你当月的多份账单（银行/信用卡、微信、支付宝）交给 AI，自动**去重对账、核定消费口径、分六大类、挑出可优化的省钱项**，产出一个 4-Sheet 的 Excel 账本 + 消费报告。核心是**口径一致、月月可比**，以及**拿不准就问、不瞎猜**。

## ⚙️ 需要什么
- 一个**能执行 Python 代码**的 AI（如 Claude Code、claude.ai 的分析工具、ChatGPT 代码解释器等）—— 因为要跑 `scripts/build_report.py` 生成 Excel。
- Python 环境装了 `openpyxl`（`pip install openpyxl`）。
- 纯聊天、不能跑代码的 AI：也能用这套**规则**帮你把账理清、口头给报告，但生成不了那个 Excel 文件。

## 🚀 怎么用
1. 把当月账单放进一个文件夹——**格式不限**：微信/支付宝的 CSV 或 Excel、信用卡/银行流水的 PDF、甚至美团月付/京东白条那种只有截图的都行（截图交给能识图的 AI，或自己把那几笔转录一下）。
2. 把 `SKILL.md` 交给你的 AI（Claude Code 放进 skill 目录；其它 AI 直接把 SKILL.md 内容发给它当指令）。
3. 说"帮我整理 X 月账单"，AI 会按 SKILL.md 的规则去重、分类、遇到拿不准的问你。
4. 你答完，它调用 `scripts/build_report.py` 生成 Excel 报告。

## 📂 目录
- `SKILL.md` — 给 AI 的完整指令（去重/消费口径/六大类/可优化项规则）
- `scripts/build_report.py` — 把整理好的数据生成 4-Sheet Excel（只依赖 openpyxl）
- `references/` — 输入字段说明 + 一份**虚构**示范数据

## ⚠️ 隐私
`references/` 里的示范数据全部虚构。**你的真实账单请放在本仓库之外、加入 `.gitignore`，切勿提交。**

## License：MIT
