# -*- coding: utf-8 -*-
"""
复盘报告生成 —— 评价周报 Word 生成器（走势 + 根因 + 对比 + 改进项）。

职责边界：本脚本只做**指标计算、基准对照、对比归因与 Word/Excel 产物生成**（机器的强项）。
根因背后的业务解释、改进项的落地资源协调由模型按 prompt.txt 完成。

周报结构（固定六段，缺数据段位标「数据缺失」不编造）：
  ① 本周评分走势  ② 指标对照（上周 vs 本周）  ③ 差评根因分布
  ④ 回复率/回复时长  ⑤ 结论与下周改进 3 项  ⑥ 人工确认事项

健康度基准（量化）：
  店铺评分 ≥ 4.7/5 达标；好评率 ≥ 95%；差评率 > 3% 须专项治理；
  回复率 100%（每条 30-80 字）；差评响应 ≤ 2 小时；一般评价 ≤ 24 小时
改进项规则：取根因分布 Top 3，每项必须带数字目标与责任人占位（人工指派）。

用法：
  python report_gen.py --input input.json --outdir out
  python report_gen.py --demo

产物：
  out/评价周报.docx       六段式 Word 周报（含内嵌图）
  out/评价周报数据.xlsx   指标对照 / 根因分布 / 改进项
  out/report.json         机器可读结果
"""
from __future__ import annotations

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
REPO = os.path.dirname(os.path.dirname(SKILL_DIR))
sys.path.insert(0, os.path.join(REPO, "lib"))

try:
    import assettools as at
except ImportError:  # pragma: no cover
    print("[错误] 未找到 lib/assettools.py。请确认技能位于 <repo>/skills/<slug>/scripts/ 下，"
          "且 <repo>/lib/assettools.py 存在。", file=sys.stderr)
    sys.exit(2)

SCORE_TARGET = 4.7       # 店铺评分达标线
GOOD_FLOOR = 0.95        # 好评率下限
BAD_CEIL = 0.03          # 差评率上限（>3% 专项治理）
REPLY_TARGET = 1.00      # 回复率 100%
REPLY_HOURS_BAD = 2      # 差评响应时限（小时）
REPLY_HOURS_NORM = 24    # 一般评价响应时限（小时）

IMPROVE_BY_CAUSE = {
    "质量问题(安全)": "下架留证批次自查，涉事订单 2 小时内全响应；品控出整改报告",
    "质量问题(工艺)": "供应商整改单 + 破损率降至 0.5% 以下；包装加缓冲",
    "物流问题": "发货 P90 压到 48 小时内；漏发率 < 0.3%",
    "服务问题": "差评会话 100% 质检；退款响应 ≤ 2 小时",
    "描述不符": "详情页逐项修正，色差类补实拍图",
    "性价比": "赠品/规格分层替代直接降价，补克重对比口径",
    "期望差(主观)": "详情页加「预期说明」，回复客观呈现口径",
}

DEMO = {
    "shop": "悦己家居旗舰店",
    "platform": "京东",
    "period": "2026-09-22 ~ 2026-09-28",
    "this_week": {
        "avg_score": 4.62, "total_reviews": 328, "good_rate": 0.936, "bad_rate": 0.046,
        "reply_rate": 0.94, "reply_hours_bad": 3.5, "reply_hours_norm": 21.0,
    },
    "last_week": {
        "avg_score": 4.68, "total_reviews": 301, "good_rate": 0.951, "bad_rate": 0.037,
        "reply_rate": 0.97, "reply_hours_bad": 2.8, "reply_hours_norm": 18.0,
    },
    "score_trend": [4.71, 4.70, 4.68, 4.65, 4.60, 4.58, 4.62],
    "root_causes": [
        {"cause": "物流问题", "count": 9},
        {"cause": "描述不符", "count": 4},
        {"cause": "质量问题(工艺)", "count": 2},
        {"cause": "服务问题", "count": 1},
    ],
}


def judge(name, cur, prev, target, better, unit=""):
    """单指标判定行：better='up' 越高越好 / 'down' 越低越好。"""
    chg = (cur - prev) if prev is not None else None
    chg_s = f"{chg:+.2f}" if isinstance(cur, float) and prev is not None else \
            (f"{chg:+.0%}" if prev is not None and isinstance(cur, (int, float)) else "—")
    if better == "up":
        ok = cur >= target
        trend = "↑" if (chg or 0) > 0 else ("↓" if (chg or 0) < 0 else "→")
    else:
        ok = cur <= target
        trend = "↓" if (chg or 0) < 0 else ("↑" if (chg or 0) > 0 else "→")
    return {"指标": name, "本周": f"{cur}{unit}" if unit else cur,
            "上周": f"{prev}{unit}" if unit and prev is not None else (prev if prev is not None else "—"),
            "变化": f"{trend} {chg_s}", "基准": target, "判定": "达标" if ok else "未达标"}


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    period = payload.get("period", "")
    this = payload.get("this_week", {})
    last = payload.get("last_week", {})
    trend = payload.get("score_trend", [])
    causes = sorted(payload.get("root_causes", []), key=lambda c: -c.get("count", 0))

    rows = [
        judge("店铺评分", this.get("avg_score", 0), last.get("avg_score"), SCORE_TARGET, "up"),
        judge("好评率", this.get("good_rate", 0), last.get("good_rate"), GOOD_FLOOR, "up", "%"),
        judge("差评率", this.get("bad_rate", 0), last.get("bad_rate"), BAD_CEIL, "down", "%"),
        judge("回复率", this.get("reply_rate", 0), last.get("reply_rate"), REPLY_TARGET, "up", "%"),
        judge("差评响应(小时)", this.get("reply_hours_bad", 0), last.get("reply_hours_bad"), REPLY_HOURS_BAD, "down"),
        judge("一般评价响应(小时)", this.get("reply_hours_norm", 0), last.get("reply_hours_norm"), REPLY_HOURS_NORM, "down"),
    ]

    # 改进项：根因 Top 3，每项带数字目标
    total_c = sum(c.get("count", 0) for c in causes) or 1
    improves = []
    for i, c in enumerate(causes[:3], 1):
        improves.append({
            "优先级": f"改进{i}", "根因": c.get("cause", "—"),
            "占比": f"{c.get('count', 0) / total_c:.0%}",
            "动作与数字目标": IMPROVE_BY_CAUSE.get(c.get("cause", ""), "人工补充动作与目标"),
            "责任人": "（周会指派）", "完成时点": "下周五",
        })

    score_now = this.get("avg_score", 0)
    head_verdict = "达标" if score_now >= SCORE_TARGET else f"低于 {SCORE_TARGET} 达标线，须专项治理"
    worst = rows[0]["判定"] == "未达标" or score_now < SCORE_TARGET
    conclusion = (f"本周评分 {score_now}，{head_verdict}。首要根因「{causes[0]['cause'] if causes else '—'}」"
                  f"占差评 {causes[0]['count'] / total_c:.0%}（数据以脚本统计为准），"
                  + ("多项基准未达标，周会逐项定责任人与数字目标。" if worst else "整体健康，改进项按周会排期。"))

    summary = {
        "店铺": shop, "平台": platform, "周期": period,
        "本周评分": score_now, "评分判定": head_verdict,
        "首要根因": causes[0]["cause"] if causes else "—",
        "改进项数": len(improves),
        "基准": {"评分": SCORE_TARGET, "好评率": GOOD_FLOOR, "差评率上限": BAD_CEIL,
                 "回复率": REPLY_TARGET, "差评响应(小时)": REPLY_HOURS_BAD},
    }

    at.ensure_outdir(outdir)
    chart = None
    if trend:
        chart = at.line_chart(
            os.path.join(outdir, "评分走势.png"),
            [f"周{i+1}" if len(trend) <= 4 else f"第{i+1}天" for i in range(len(trend))],
            {"店铺评分": trend}, title=f"{period} 评分走势", ylabel="评分",
        )

    xlsx = at.write_excel(
        os.path.join(outdir, "评价周报数据.xlsx"),
        {
            "指标对照": rows,
            "根因分布": [{"根因": c.get("cause"), "条数": c.get("count"),
                          "占比": f"{c.get('count', 0) / total_c:.0%}"} for c in causes],
            "改进项": improves,
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"指标对照": {"判定": "contains:未达标"}},
        widths={"改进项": {"动作与数字目标": 48}},
    )

    sections = [
        {"heading": "一、核心结论", "paras": [conclusion]},
        {"heading": "二、本周评分走势", "paras": [f"评分走势见下图；本周均值 {score_now}，达标线 {SCORE_TARGET}。"],
         "image": chart},
        {"heading": "三、指标对照（本周 vs 上周）",
         "table": {"cols": ["指标", "本周", "上周", "变化", "基准", "判定"],
                   "rows": [[r["指标"], r["本周"], r["上周"], r["变化"], r["基准"], r["判定"]] for r in rows]}},
        {"heading": "四、差评根因分布",
         "table": {"cols": ["根因", "条数", "占比"],
                   "rows": [[c.get("cause"), c.get("count"), f"{c.get('count', 0) / total_c:.0%}"] for c in causes]}},
        {"heading": "五、回复情况",
         "paras": [f"回复率 {this.get('reply_rate', 0):.0%}（目标 100%，每条 30-80 字个性化回复）；"
                    f"差评响应 {this.get('reply_hours_bad', '—')} 小时（目标 ≤ {REPLY_HOURS_BAD} 小时）；"
                    f"一般评价响应 {this.get('reply_hours_norm', '—')} 小时（目标 ≤ {REPLY_HOURS_NORM} 小时）。"]},
        {"heading": "六、下周改进 3 项（周会定责任人）",
         "table": {"cols": ["优先级", "根因", "占比", "动作与数字目标", "责任人", "完成时点"],
                   "rows": [[r["优先级"], r["根因"], r["占比"], r["动作与数字目标"], r["责任人"], r["完成时点"]] for r in improves]}},
        {"heading": "七、人工确认事项",
         "bullets": ["改进项责任人与完成时点需周会确认；涉诉差评处置前人工复核",
                     "本报告为 AI 生成内容，数据来自平台导出，未经人工核对不得对外引用"]},
    ]
    docx = at.write_docx(
        os.path.join(outdir, "评价周报.docx"),
        f"{shop} 评价周报（{period}）",
        sections, subtitle=f"{platform} · AI 生成内容 · 数值以脚本统计为准",
    )

    js = at.write_json({"summary": summary, "metrics": rows, "improves": improves,
                        "generated_at": at.stamp(),
                        "note": "周报数据；根因业务解释由模型按 prompt.txt 补充"},
                       os.path.join(outdir, "report.json"))
    files = [f for f in (docx, xlsx, chart, js) if f]
    return {"files": files, "summary": summary, "conclusion": conclusion}


def main():
    ap = argparse.ArgumentParser(description="复盘报告生成 —— 评价周报 Word")
    ap.add_argument("--input", help="输入 JSON（shop/this_week/last_week/score_trend/root_causes）")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--demo", action="store_true")
    a = ap.parse_args()

    if a.demo:
        payload = DEMO
    elif a.input:
        payload = at.read_json(a.input)
    else:
        ap.error("需要 --input / --demo 之一")

    r = build(payload, a.outdir)
    print(r["conclusion"])
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
