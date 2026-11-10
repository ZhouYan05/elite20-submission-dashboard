# -*- coding: utf-8 -*-
"""C6A 报告生成器：submissions.json + challenges.json -> CSV + Excel 五表 + 仪表盘 data.json
用法: python gen_report.py [数据目录]
"""
import csv
import json
import os
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))

STATUS_STYLE = {
    "checked":  ("🔎 已提交·平台校验通过", "FFF2CC"),
    "complete": ("✅ 完整", "C6EFCE"),
    "partial":  ("⚠️ 部分/待修", "FFEB9C"),
    "submitted": ("📤 已提交·待校验", "DDEBF7"),
    "missing":  ("❌ 未提交", "FFC7CE"),
    "skipped":  ("⏭️ 计划不提交", "D9D9D9"),
}


def load(name):
    with open(os.path.join(DATA_DIR, name), encoding="utf-8") as f:
        return json.load(f)


def main():
    subs = load("submissions.json")
    chals = load("challenges.json")
    by_id = {c["id"]: c for c in chals["challenges"]}

    # ---- 1) 合并后的仪表盘数据 ----
    dashboard = {
        "meta": subs["meta"],
        "challenges": chals["challenges"],
        "submissions": subs["submissions"],
        "stats": {},
    }
    # 完成率统计（排除 T0）
    active = [s for s in subs["submissions"] if s["challenge"] != "T0"]
    done = [s for s in active if s["status"] in ("checked", "complete", "submitted")]
    partial = [s for s in active if s["status"] == "partial"]
    missing = [s for s in active if s["status"] in ("missing", "skipped")]
    scored = [s for s in active if s.get("score")]
    dashboard["stats"] = {
        "total_challenges": len(active),
        "submitted": len(done),
        "partial": len(partial),
        "missing": len(missing),
        "completion_rate": round(len(done) / len(active) * 100, 1) if active else 0,
        "avg_score": round(sum(s["score"] for s in scored) / len(scored), 1) if scored else None,
        "scored_count": len(scored),
    }
    stats = dashboard["stats"]
    with open(os.path.join(DATA_DIR, "dashboard_data.json"), "w", encoding="utf-8") as f:
        json.dump(dashboard, f, ensure_ascii=False, indent=2)

    # 仪表盘内嵌数据（file:// 与 GitHub Pages 均可直接加载）
    dash_js = os.path.join(DATA_DIR, "..", "ZhouYan_C6A_仪表盘", "data.js")
    os.makedirs(os.path.dirname(dash_js), exist_ok=True)
    with open(dash_js, "w", encoding="utf-8") as f:
        f.write("window.DASHBOARD_DATA = " + json.dumps(dashboard, ensure_ascii=False) + ";\n")

    # 每周提交状态摘要（发群通知用 Markdown）
    summary_lines = [
        "# Elite20 提交状态周报（自动生成）",
        "",
        f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"- 学生：{subs['meta']['student']['name']}（{subs['meta']['student']['github']}）",
        "",
        "## 总览",
        "",
        f"- 有效挑战 **{stats['total_challenges']}** 个：已提交 **{stats['submitted']}**、部分待修 **{stats['partial']}**、"
        f"未提交 **{stats['missing']}**；完成率 **{stats['completion_rate']}%**",
    ]
    if stats["avg_score"] is not None:
        summary_lines.append(f"- 已获 AI 评分：{stats['scored_count']} 项，平均 {stats['avg_score']} 分")
    summary_lines += ["", "## 需要关注", ""]
    for s in subs["submissions"]:
        if s["status"] in ("partial", "missing"):
            summary_lines.append(f"- **{s['challenge']}**：{s.get('notes') or '未提交'}")
    summary_lines += ["", "## 最近提交", ""]
    recent = [s for s in subs["submissions"] if s.get("submitted_at") and s["challenge"] != "T0"]
    recent.sort(key=lambda s: str(s["submitted_at"]), reverse=True)
    for s in recent[:5]:
        summary_lines.append(f"- {s['challenge']}：{s.get('project_title')}（{s['submitted_at'][:10]}）")
    with open(os.path.join(DATA_DIR, "weekly_summary.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(summary_lines) + "\n")

    # ---- 2) CSV ----
    with open(os.path.join(DATA_DIR, "submissions.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["挑战", "标题", "学生", "提交时间", "状态", "评分", "仓库", "备注"])
        for s in subs["submissions"]:
            c = by_id.get(s["challenge"], {})
            w.writerow([
                s["challenge"], c.get("title", ""), subs["meta"]["student"]["name"],
                s.get("submitted_at") or "", s["status"], s.get("score") or "",
                s.get("github_repo") or "", s.get("notes") or "",
            ])

    # ---- 3) Excel 五表 ----
    wb = Workbook()

    def style_header(ws, ncols):
        fill = PatternFill("solid", fgColor="4472C4")
        for i in range(1, ncols + 1):
            c = ws.cell(row=1, column=i)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = fill
            c.alignment = Alignment(horizontal="center")

    # Sheet1 总览：学生 × 挑战 矩阵
    ws = wb.active
    ws.title = "总览"
    ids = [c["id"] for c in chals["challenges"]]
    ws.append(["学生"] + ids)
    row = [subs["meta"]["student"]["name"]]
    for cid in ids:
        s = next((x for x in subs["submissions"] if x["challenge"] == cid), None)
        label, _ = STATUS_STYLE.get(s["status"] if s else "missing", STATUS_STYLE["missing"])
        row.append(label.split(" ")[0])
    ws.append(row)
    stats = dashboard["stats"]
    ws.append(["完成率统计", f"已提交 {stats['submitted']} / 共 {stats['total_challenges']} 个有效挑战，"
                            f"完成率 {stats['completion_rate']}%（部分待修 {stats['partial']}，未提交 {stats['missing']}）"])
    style_header(ws, len(ids) + 1)
    for r in ws.iter_rows(min_row=2, max_row=2, min_col=2, max_col=len(ids) + 1):
        for c in r:
            s = next((x for x in subs["submissions"] if x["challenge"] == c.column_letter and False), None)
            c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 12

    # Sheet2 按学生：详情
    ws2 = wb.create_sheet("按学生")
    ws2.append(["挑战", "标题", "项目名", "提交时间", "状态", "AI评分", "交付文件", "备注"])
    for s in subs["submissions"]:
        c = by_id.get(s["challenge"], {})
        label, _ = STATUS_STYLE.get(s["status"], ("", ""))
        ws2.append([s["challenge"], c.get("title", ""), s.get("project_title") or "",
                    s.get("submitted_at") or "", label, s.get("score") or "",
                    "；".join(s.get("files") or []), s.get("notes") or ""])
    style_header(ws2, 8)
    widths = [7, 30, 42, 20, 20, 9, 46, 40]
    for i, wd in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = wd

    # Sheet3 按挑战：全部挑战的提交情况（含必填交付物模式）
    ws3 = wb.create_sheet("按挑战")
    ws3.append(["挑战", "标题", "必填交付物模式", "提交状态", "提交时间", "仓库", "评分"])
    for c in chals["challenges"]:
        s = next((x for x in subs["submissions"] if x["challenge"] == c["id"]), None)
        label, _ = STATUS_STYLE.get(s["status"] if s else "missing", ("", ""))
        ws3.append([c["id"], c["title"], c.get("deliverables_pattern") or "—",
                    label, (s.get("submitted_at") or "") if s else "",
                    (s.get("github_repo") or "") if s else "", (s.get("score") or "") if s else ""])
    style_header(ws3, 7)
    for i, wd in enumerate([7, 34, 42, 20, 20, 44, 8], 1):
        ws3.column_dimensions[get_column_letter(i)].width = wd

    # Sheet4 时间线
    ws4 = wb.create_sheet("时间线")
    ws4.append(["提交时间", "挑战", "项目名", "状态", "评分"])
    ordered = [s for s in subs["submissions"] if s.get("submitted_at")]
    ordered.sort(key=lambda s: str(s.get("submitted_at")))
    for s in ordered:
        label, _ = STATUS_STYLE.get(s["status"], ("", ""))
        ws4.append([s["submitted_at"], s["challenge"], s.get("project_title") or "",
                    label, s.get("score") or ""])
    style_header(ws4, 5)
    for i, wd in enumerate([22, 8, 55, 20, 8], 1):
        ws4.column_dimensions[get_column_letter(i)].width = wd

    # Sheet5 缺失项：标红待修/缺交付物的条目
    ws5 = wb.create_sheet("缺失项")
    ws5.append(["挑战", "问题类型", "说明", "建议动作"])
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    yellow_fill = PatternFill("solid", fgColor="FFEB9C")
    for s in subs["submissions"]:
        c = by_id.get(s["challenge"], {})
        if s["status"] == "partial":
            ws5.append([s["challenge"], "平台退回 needs_revision", s.get("notes") or "",
                        "按平台反馈补交付物后重提"])
            for cell in ws5[ws5.max_row]:
                cell.fill = red_fill
        elif s["status"] == "skipped" and s["challenge"] != "T0":
            ws5.append([s["challenge"], "本轮未提交", s.get("notes") or "", "按个人计划决定是否补做"])
            for cell in ws5[ws5.max_row]:
                cell.fill = yellow_fill
    style_header(ws5, 4)
    for i, wd in enumerate([7, 24, 50, 32], 1):
        ws5.column_dimensions[get_column_letter(i)].width = wd

    out = os.path.join(DATA_DIR, "..", "ZhouYan_C6A_report.xlsx")
    wb.save(out)
    print("OK ->", os.path.abspath(out))
    print("stats:", json.dumps(dashboard["stats"], ensure_ascii=False))


if __name__ == "__main__":
    main()
