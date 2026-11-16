# Demo 说明（C6A 演示证据）

## 演示入口

- 在线仪表盘：**<https://ZhouYan05.github.io/elite20-submission-dashboard/>**
- 仓库：<https://github.com/ZhouYan05/elite20-submission-dashboard>
- 数据管线入口脚本：`data/gen_report.py`

## 一条完整演示链（约 2 分钟）

1. **打开在线仪表盘** → 顶部总览卡片显示：有效挑战数、已提交数、完成率、平均 AI 评分。
2. **按状态筛选** → 选 `needs_revision`，列表只剩被平台退回的挑战（例如 C4A / C4B），点开可见退回原因与需补项。
3. **按分数区间筛选 / 排序** → 按 score_total 降序，最快看到「分数低且未闭环」的组合。
4. **查看单条详情** → 挑战名、提交时间、状态、评分、评审反馈摘要。
5. **导出报表** → 仓库内 `report.xlsx`（由 `data/gen_report.py` 从 `data/submissions.json` + `data/challenges.json` 生成，5 个工作表）。

## 数据流演示

```
data/submissions.json  ┐
data/challenges.json   ├→ gen_report.py → submissions.csv
                       │                → report.xlsx（5 sheets）
                       │                → dashboard_data.json → data.js
                       │                → weekly_summary.md（周报摘要）
```

复现命令：

```bash
cd elite20-submission-dashboard/data
pip install openpyxl
python gen_report.py
```

## 演示覆盖的能力点

| 能力点 | 演示位置 |
|---|---|
| 真实数据驱动（非手写假数据） | `data/submissions.json` 来自平台 `list-my-submissions` 导出 |
| 交互式浏览（筛选/排序/详情） | 在线仪表盘页面 |
| 加工产物可交付 | `report.xlsx` 五表 + `weekly_summary.md` |
| 管线可复现 | `gen_report.py` 一条命令重建全部产物 |
| 结论可用（不只是图表） | `weekly_summary.md` 里的「需要关注」直接指向 C4A/C4B 待修 |
