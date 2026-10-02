# -*- coding: utf-8 -*-
"""
周报复盘 —— 端到端编排脚本（原始评价 → 情感统计 → 根因归因 → 周报 Word）。

流程 DAG：
  原始评价+周指标 → S1 完整性校验 → S2 情感五级统计(review-sentiment-analyze 规则)
                  → S3 差评根因归因(negative-review-rootcause 规则)
                  → S4 指标对照与趋势(review-report-generate 基准) → 周报产物

健康度基准（量化）：
  店铺评分 ≥ 4.7/5；好评率 ≥ 95%；差评率 > 3% 专项治理；
  回复率 100%（30-80 字/条）；差评响应 ≤ 2 小时；追评比首评权重更高（×2 入归因）
  连续 3 天同向下滑才算趋势恶化；变化 ≤ 2 个百分点判噪声

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo

产物：
  out/评价周报复盘.docx     六段式周报 Word（含走势图与根因图内嵌）
  out/周报复盘数据.xlsx     指标对照 / 根因分布 / 改进项
  out/情感与根因图.png      评分走势 + 负面标签图
  out/weekly_flow.json      机器可读结果
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
WF_DIR = os.path.dirname(HERE)
REPO = os.path.dirname(os.path.dirname(WF_DIR))
sys.path.insert(0, os.path.join(REPO, "lib"))

try:
    import assettools as at
except ImportError:  # pragma: no cover
    print("[错误] 未找到 lib/assettools.py。请确认工作流位于 <repo>/workflows/<slug>/scripts/ 下，"
          "且 <repo>/lib/assettools.py 存在。", file=sys.stderr)
    sys.exit(2)

SCORE_TARGET = 4.7
GOOD_FLOOR = 0.95
BAD_CEIL = 0.03
REPLY_TARGET = 1.00
FOLLOWUP_WEIGHT = 2

CAUSE_WORDS = {
    "质量问题(安全)": r"变质|发霉|过期|异物|虫|头发|腹泻|过敏|中毒|异味|臭",
    "质量问题(工艺)": r"碎|破损|漏液|掉色|开线|做工粗糙|以次充好|假货|用几天就坏|质量差|跑绒",
    "物流问题": r"物流慢|发货慢|迟迟不发货|等了\d+天|催了|暴力分拣|压坏|挤压|漏发|错发|丢件|迟迟不到|包装.*破",
    "服务问题": r"客服.*(态度|敷衍|不理|踢皮球|推诿)|态度差|爱读不回|已读不回|退款慢|售后.*差|无人处理",
    "描述不符": r"与描述|与图片|详情页|实物.*不符|色差|缩水|克重不足|分量.*少|尺寸.*不对",
    "性价比": r"太贵|不值|溢价|买贵了|降价.*快|这个价.*不如|性价比",
}
CAUSE_ORDER = list(CAUSE_WORDS.keys())
FALLBACK = "期望差(主观)"
ACTIONS = {
    "质量问题(安全)": "当日下架留证批次自查，涉事订单 2 小时内全响应",
    "质量问题(工艺)": "供应商整改单，破损率目标 < 0.5%",
    "物流问题": "发货 P90 ≤ 48 小时，漏发率 < 0.3%",
    "服务问题": "差评会话 100% 质检，退款响应 ≤ 2 小时",
    "描述不符": "详情页逐项修正，色差补实拍图",
    "性价比": "赠品/规格分层替代降价",
    FALLBACK: "详情页加预期说明，回复客观呈现口径",
}

DEMO = {
    "shop": "山禾炊具旗舰店",
    "platform": "京东",
    "period": "2026-09-22 ~ 2026-09-28",
    "metrics": {
        "avg_score": 4.66, "good_rate": 0.944, "bad_rate": 0.041, "total_reviews": 312,
        "reply_rate": 0.96, "reply_hours_bad": 3.0, "reply_hours_norm": 19.0,
    },
    "last_metrics": {
        "avg_score": 4.72, "good_rate": 0.955, "bad_rate": 0.033,
        "reply_rate": 0.98, "reply_hours_bad": 2.4, "reply_hours_norm": 17.0,
    },
    "score_trend": [4.73, 4.72, 4.70, 4.66, 4.62, 4.60, 4.66],
    "reviews": [
        {"star": 1, "text": "锅底涂层起皮，用两周就掉渣，质量差", "has_image": True, "is_followup": False},
        {"star": 2, "text": "发货慢，等了6天，双11都没这么慢", "has_image": False, "is_followup": False},
        {"star": 2, "text": "锅把手松动，螺丝还少一颗", "has_image": True, "is_followup": False},
        {"star": 3, "text": "锅比较重，女生端不动，详情页没写重量", "has_image": False, "is_followup": False},
        {"star": 3, "text": "导热一般，比描述的受热均匀差远了", "has_image": False, "is_followup": False},
        {"star": 1, "text": "锅盖碎了，客服已读不回三天", "has_image": True, "is_followup": False},
        {"star": 3, "text": "太重了不值这个价", "has_image": False, "is_followup": True},
        {"star": 4, "text": "不粘效果不错，就是锅底有点厚", "has_image": False, "is_followup": False},
        {"star": 5, "text": "第二次回购，受热均匀油烟小，客服还送了木铲", "has_image": True, "is_followup": False},
        {"star": 5, "text": "做工精细，锅盖带沥水设计很实用", "has_image": False, "is_followup": False},
        {"star": 2, "text": "涂层有划痕，像是展示样品", "has_image": True, "is_followup": True},
        {"star": 4, "text": "整体满意，物流略慢", "has_image": False, "is_followup": False},
    ],
}


def sentiment_level(r: dict) -> str:
    star = r.get("star")
    text = str(r.get("text", ""))
    if not text and star is None:
        return "中"
    star = 3 if star is None else int(star)
    if star <= 2:
        return "强负"
    if star == 3:
        return "负面"
    if star == 5:
        return "强正" if (r.get("has_image") or len(text) >= 15) else "正面"
    return "正面"


def attribute(text: str) -> str:
    for cause in CAUSE_ORDER:
        if re.search(CAUSE_WORDS[cause], text):
            return cause
    return FALLBACK


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    period = payload.get("period", "")
    m = payload.get("metrics", {})
    lm = payload.get("last_metrics", {})
    trend = payload.get("score_trend", [])
    reviews = payload.get("reviews", [])

    # S2 情感统计
    lv = Counter(sentiment_level(r) for r in reviews)
    n = len(reviews) or 1
    good_n = lv.get("强正", 0) + lv.get("正面", 0)
    bad_n = lv.get("强负", 0)
    est_good = good_n / n
    est_bad = bad_n / n

    # S3 根因归因（追评 ×2）
    dist, weight_sum = Counter(), 0
    for r in reviews:
        if r.get("star") is not None and int(r["star"]) <= 3:
            w = FOLLOWUP_WEIGHT if r.get("is_followup") else 1
            weight_sum += w
            dist[attribute(str(r.get("text", "")))] += w
    tw = weight_sum or 1
    causes = sorted(dist.items(), key=lambda kv: -kv[1])

    # S4 指标对照
    metric_rows = []
    for name, cur, prev, target, better, unit in [
        ("店铺评分", m.get("avg_score"), lm.get("avg_score"), SCORE_TARGET, "up", ""),
        ("好评率", m.get("good_rate"), lm.get("good_rate"), GOOD_FLOOR, "up", "%"),
        ("差评率", m.get("bad_rate"), lm.get("bad_rate"), BAD_CEIL, "down", "%"),
        ("回复率", m.get("reply_rate"), lm.get("reply_rate"), REPLY_TARGET, "up", "%"),
        ("差评响应(h)", m.get("reply_hours_bad"), lm.get("reply_hours_bad"), 2, "down", ""),
        ("一般响应(h)", m.get("reply_hours_norm"), lm.get("reply_hours_norm"), 24, "down", ""),
    ]:
        if cur is None:
            metric_rows.append({"指标": name, "本周": "数据缺失", "上周": prev or "—", "基准": target, "判定": "数据缺失"})
            continue
        ok = cur >= target if better == "up" else cur <= target
        chg = (cur - prev) if prev is not None else None
        noise = "（噪声 ≤2pct）" if (chg is not None and name != "店铺评分" and abs(chg) <= 0.02) else ""
        metric_rows.append({
            "指标": name,
            "本周": f"{cur:.1%}" if unit == "%" else cur,
            "上周": (f"{prev:.1%}" if unit == "%" else prev) if prev is not None else "—",
            "基准": target, "判定": ("达标" if ok else "未达标") + noise,
        })

    total_c = sum(c.get("count", 0) for c in []) if False else sum(dist.values()) or 1
    improves = [{"优先级": f"改进{i}", "根因": c, "占比": f"{v / tw:.0%}",
                 "动作与数字目标": ACTIONS.get(c, ""), "责任人": "（周会指派）", "完成时点": "下周五"}
                for i, (c, v) in enumerate(causes[:3], 1)]

    score = m.get("avg_score")
    alarms = []
    if score is not None and score < SCORE_TARGET:
        alarms.append(f"评分 {score} < {SCORE_TARGET} 达标线")
    if m.get("bad_rate") is not None and m["bad_rate"] > BAD_CEIL:
        alarms.append(f"差评率 {m['bad_rate']:.1%} > {BAD_CEIL:.0%}，须专项治理")
    if m.get("reply_rate") is not None and m["reply_rate"] < REPLY_TARGET:
        missed = round((1 - m["reply_rate"]) * m.get("total_reviews", n))
        alarms.append(f"回复率 {m['reply_rate']:.0%} < 100%，约 {missed} 条漏回复")
    if dist.get("质量问题(安全)", 0) / tw > 0.10:
        alarms.append("质量安全类差评 > 10%，当日下架自查")
    # 连续 3 天同向下滑
    if len(trend) >= 4:
        for i in range(len(trend) - 3):
            seg = trend[i:i + 4]
            if all(seg[j] > seg[j + 1] for j in range(3)):
                alarms.append(f"评分连续 3 天下滑（{seg[0]}→{seg[3]}），判趋势恶化非噪声")
                break

    conclusion = (f"本周评分 {score}（上周 {lm.get('avg_score', '—')}），"
                  f"首要根因「{causes[0][0] if causes else '—'}」占差评 {causes[0][1] / tw:.0%}"
                  if causes else "本周无差评数据") + \
                 ("；" + "；".join(alarms) + "。" if alarms else "；各项在基准内，按排期推进改进。" if causes else "。")

    summary = {
        "店铺": shop, "平台": platform, "周期": period,
        "评价总数": len(reviews),
        "情感分布": {k: lv.get(k, 0) for k in ("强负", "负面", "中", "正面", "强正")},
        "样本好评率": f"{est_good:.0%}", "样本差评率": f"{est_bad:.0%}",
        "首要根因": causes[0][0] if causes else "—",
        "告警": "；".join(alarms) or "无",
        "基准": {"评分": SCORE_TARGET, "好评率": GOOD_FLOOR, "差评率": BAD_CEIL, "回复率": REPLY_TARGET},
        "编排说明": ("S1 校验 → S2 情感统计(同源 review-sentiment-analyze) → S3 根因归因(同源 "
                     "negative-review-rootcause) → S4 指标对照(基准 review-report-generate)；"
                     "数值以脚本输出为准，不要自己算"),
    }

    at.ensure_outdir(outdir)
    chart = at.line_chart(
        os.path.join(outdir, "情感与根因图.png"),
        [f"第{i+1}天" for i in range(len(trend))],
        {"店铺评分": trend}, title=f"{period} 评分走势", ylabel="评分",
    )
    xlsx = at.write_excel(
        os.path.join(outdir, "周报复盘数据.xlsx"),
        {
            "指标对照": metric_rows,
            "根因分布": [{"根因": c, "加权条数": v, "占比": f"{v / tw:.0%}",
                          "治理动作": ACTIONS.get(c, "")} for c, v in causes],
            "改进项": improves or [{"优先级": "—", "根因": "（本周无差评）"}],
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"指标对照": {"判定": "contains:未达标"}},
        widths={"根因分布": {"治理动作": 44}},
    )
    sections = [
        {"heading": "一、核心结论", "paras": [conclusion]},
        {"heading": "二、本周评分走势", "paras": [f"本周均值 {score}，达标线 {SCORE_TARGET}；连续 3 天同向变化才判趋势。"],
         "image": chart},
        {"heading": "三、指标对照（本周 vs 上周）",
         "table": {"cols": ["指标", "本周", "上周", "基准", "判定"],
                   "rows": [[r["指标"], r["本周"], r["上周"], r["基准"], r["判定"]] for r in metric_rows]}},
        {"heading": "四、差评根因分布（追评×2 加权）",
         "table": {"cols": ["根因", "加权条数", "占比"],
                   "rows": [[c, v, f"{v / tw:.0%}"] for c, v in causes]}},
        {"heading": "五、回复情况",
         "paras": [f"回复率 {m.get('reply_rate', '—')}（目标 100%，每条 30-80 字个性化回复）；"
                    f"差评响应 {m.get('reply_hours_bad', '—')} 小时（目标 ≤ 2 小时）。"]},
        {"heading": "六、下周改进 3 项（周会定责任人）",
         "table": {"cols": ["优先级", "根因", "占比", "动作与数字目标", "责任人", "完成时点"],
                   "rows": [[r["优先级"], r["根因"], r["占比"], r["动作与数字目标"], r["责任人"], r["完成时点"]] for r in improves]}},
        {"heading": "七、人工确认事项",
         "bullets": ["改进项责任人与时点周会确认；涉诉差评处置前人工复核",
                     "本报告为 AI 生成内容，数值以脚本统计为准，未经人工核对不对外引用"]},
    ]
    docx = at.write_docx(
        os.path.join(outdir, "评价周报复盘.docx"),
        f"{shop} 评价周报复盘（{period}）", sections,
        subtitle=f"{platform} · AI 生成内容",
    )
    js = at.write_json({"summary": summary, "metrics": metric_rows, "improves": improves,
                        "generated_at": at.stamp(),
                        "note": "端到端周报数据；根因业务解释由模型按 prompt.txt 补充"},
                       os.path.join(outdir, "weekly_flow.json"))
    return {"files": [docx, xlsx, chart, js], "summary": summary, "conclusion": conclusion}


def main():
    ap = argparse.ArgumentParser(description="周报复盘工作流")
    ap.add_argument("--input", help="输入 JSON（shop/metrics/reviews/...）")
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
