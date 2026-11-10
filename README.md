# elite20-submission-dashboard — Elite20 提交数据仪表盘

一句话描述：**输入 Elite20 课程的提交记录（JSON），输出可浏览、可筛选、可追踪的交互式仪表盘 + 五表 Excel 报表 + 周报摘要。**

在线查看：<https://ZhouYan05.github.io/elite20-submission-dashboard/>

## 解决什么问题

课程提交物散落在群文件、聊天记录、各人工作区与 GitHub 仓库里，没有人能一眼看清「谁交了、交了啥、缺了啥」。本项目把散落信息收敛为一条可复跑的数据管线：

```
平台提交记录 / 本地交付物 / GitHub 仓库树
        ↓ 收集核对
submissions.json（单一事实源）+ challenges.json（挑战目录）
        ↓ gen_report.py（一键重生成）
CSV · Excel 五表 · data.js（仪表盘数据）· weekly_summary.md（周报）
```

## 快速开始

### 在线看仪表盘（零安装）

<https://ZhouYan05.github.io/elite20-submission-dashboard/>

功能：完成率卡片、挑战×状态热力图、按状态筛选 / 按时间排序 / 关键词搜索、点击行查看提交详情与交付文件清单、导出 CSV、下载 Excel 报表。

### 本地打开

```bash
git clone https://github.com/ZhouYan05/elite20-submission-dashboard.git
cd elite20-submission-dashboard
# 双击 index.html 即可（数据内嵌 data.js，file:// 直接可用）
```

### 更新数据（可复跑管线）

```bash
# 1. 编辑 data/submissions.json（新增/修改提交记录）
# 2. 需要 Python 3 + openpyxl，重生成全部产物：
python data/gen_report.py data
# 产出：data/dashboard_data.json、data.js、data/submissions.csv、report.xlsx、data/weekly_summary.md
```

## 示例

输入 `data/submissions.json`：

```json
{"challenge": "C4C", "status": "checked", "score": 84,
 "project_title": "C4C 作业自动求解与排版 — …", "submitted_at": "2026-10-04T14:19:31Z", …}
```

输出：仪表盘热力图中 C4C 单元格显示 🟡（平台校验通过），详情页展示 84 分、交付文件清单与评审反馈；Excel「缺失项」表自动标红 `partial` 条目。

## 项目结构

```
elite20-submission-dashboard/
├── README.md
├── LICENSE
├── .gitignore
├── index.html          ← 仪表盘单页（零依赖，数据内嵌）
├── data.js             ← 内嵌数据（由 gen_report.py 生成）
├── data/
│   ├── submissions.json   ← 单一事实源（提交记录）
│   ├── challenges.json    ← 挑战目录（含必填交付物模式）
│   ├── dashboard_data.json
│   ├── submissions.csv
│   ├── weekly_summary.md  ← 周报摘要（发群用）
│   └── gen_report.py      ← 生成器（CSV/Excel/data.js/周报）
└── report.xlsx         ← 五表报表（总览/按学生/按挑战/时间线/缺失项）
```

## 技术栈

- 数据：JSON（单一事实源）+ CSV（交换）
- 生成器：Python 3 + openpyxl（Excel 五表）
- 前端：原生 HTML/CSS/JS，零依赖（file:// 与 GitHub Pages 均可运行）
- 部署：GitHub Pages

## AI 生成说明

本项目使用 AI 辅助开发（Claude Code），详见交付包中的 `ZhouYan_C6A_AI日志.md`。

## 借鉴来源

数据模型参考 C6A 挑战规范 Schema，详见交付包中的 `ZhouYan_C6A_拿来说明.md`。

## License

MIT
