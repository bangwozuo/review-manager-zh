# -*- coding: utf-8 -*-
"""
差评根因归因 —— 端到端编排脚本。

流程 DAG：
  差评收集 → S1 完整性与恶意差评筛查 → S2 六类根因归因(negative-review-rootcause 规则)
           → S3 趋势对比(本周 vs 上周占比，>10pct 变化才算真实) → S4 治理排期 → 汇总产物

编排的原子技能：de_ecom_03_sk02 差评根因分类（归因规则同源）。

趋势判定规则（量化）：
  占比变化 ≤ 10 个百分点 → 判噪声不下结论（小样本占比天然波动大）
  样本 < 10 条 → 不做占比对比，只逐条处置
  新增根因（上周 0 条本周 ≥ 3 条）→ 标「新发问题」优先排查（如换供应商/换快递的时间点）

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo

产物：
  out/差评归因治理表.xlsx   归因明细 / 趋势对比 / 治理排期
  out/根因趋势.png          本周 vs 上周根因占比对比图
  out/rootcause_flow.json   机器可读结果
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

FOLLOWUP_WEIGHT = 2
NOISE_BAND = 0.10        # 占比变化 ≤ 10pct 判噪声
MIN_SAMPLE = 10          # 样本 < 10 不做占比对比
MALICIOUS_PAT = r"不加微信.*不删|加V|加微信.*改|红包.*删|返.*元.*改|同行|低价.*卖.*同款"

CAUSE_WORDS = {
    "质量问题(安全)": r"变质|发霉|过期|异物|虫|头发|腹泻|过敏|中毒|异味|臭",
    "质量问题(工艺)": r"碎|破损|漏液|掉色|开线|做工粗糙|以次充好|假货|用几天就坏|质量差",
    "物流问题": r"物流慢|发货慢|迟迟不发货|等了\d+天|催了|暴力分拣|压坏|挤压|漏发|错发|丢件|迟迟不到|包装.*破",
    "服务问题": r"客服.*(态度|敷衍|不理|踢皮球|推诿)|态度差|爱答不理|已读不回|退款慢|售后.*差|无人处理|互相推",
    "描述不符": r"与描述|与图片|详情页|实物.*不符|色差|缩水|克重不足|分量.*少|尺寸.*不对|规格.*不对",
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
    "shop": "轻羽羽绒服饰店",
    "platform": "淘宝",
    "period": "2026-09-22 ~ 2026-09-28",
    "event_note": "9-20 更换薄绒供应商",
    "negatives": [
        {"star": 1, "text": "绒子刺鼻异味，晾了三天还在，怀疑不是新绒", "has_image": False, "is_followup": False},
        {"star": 2, "text": "拉链卡顿，袖口开线，做工不如去年", "has_image": True, "is_followup": False},
        {"star": 2, "text": "发货慢，等了7天，客服只会说抱歉", "has_image": False, "is_followup": False},
        {"star": 3, "text": "实物颜色和详情页差太多，明显色差", "has_image": True, "is_followup": False},
        {"star": 1, "text": "穿一周就跑绒，里面全是毛", "has_image": False, "is_followup": True},
        {"star": 2, "text": "尺码偏大两码，与描述的尺码表对不上", "has_image": False, "is_followup": False},
        {"star": 3, "text": "物流慢，半个月才到偏远的地区", "has_image": False, "is_followup": True},
        {"star": 2, "text": "太贵了，这个填充量不值", "has_image": False, "is_followup": False},
        {"star": 1, "text": "异味大得头疼，孩子穿着直挠脖子", "has_image": True, "is_followup": False},
        {"star": 2, "text": "客服推诿不退运费", "has_image": False, "is_followup": False},
        {"star": 3, "text": "版型一般，不如图片好看", "has_image": False, "is_followup": False},
        {"star": 2, "text": "开线+扣子掉了，做工太差", "has_image": False, "is_followup": True},
    ],
    "last_week_dist": {"质量问题(工艺)": 0.18, "物流问题": 0.22, "服务问题": 0.15,
                       "描述不符": 0.20, "性价比": 0.10, "期望差(主观)": 0.15},
}


def attribute(text: str):
    primary = None
    for cause in CAUSE_ORDER:
        if re.search(CAUSE_WORDS[cause], text):
            primary = cause
            break
    return primary or FALLBACK


def is_malicious(text: str) -> bool:
    return bool(re.search(MALICIOUS_PAT, text))


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    period = payload.get("period", "")
    event_note = payload.get("event_note", "")
    negatives = payload.get("negatives", [])
    last_dist = payload.get("last_week_dist", {})

    detail, dist, malicious = [], Counter(), []
    weight_sum = 0
    for i, r in enumerate(negatives, 1):
        text = str(r.get("text", ""))
        if is_malicious(text):
            malicious.append(f"#{i} {text[:24]}…")
            continue
        cause = attribute(text)
        weight = FOLLOWUP_WEIGHT if r.get("is_followup") else 1
        weight_sum += weight
        dist[cause] += weight
        detail.append({
            "序号": i, "星级": r.get("star", "—"), "根因": cause,
            "权重": f"×{weight}" if weight > 1 else "×1",
            "图证": "有" if r.get("has_image") else "无",
            "摘要": text[:28] + ("…" if len(text) > 28 else ""),
        })

    total_w = weight_sum or 1
    this_dist = {c: dist.get(c, 0) / total_w for c in list(CAUSE_WORDS) + [FALLBACK]}

    # S3 趋势对比
    trend_rows, new_issues = [], []
    for c, cur in sorted(this_dist.items(), key=lambda kv: -kv[1]):
        prev = last_dist.get(c)
        if prev is None:
            chg, verdict = "—", cur and "新增基线缺失" or "—"
        else:
            delta = cur - prev
            chg = f"{delta:+.0%}"
            verdict = "噪声（≤10pct）" if abs(delta) <= NOISE_BAND else ("↑ 恶化" if delta > 0 else "↓ 改善")
        trend_rows.append({"根因": c, "本周占比": f"{cur:.0%}",
                           "上周占比": f"{prev:.0%}" if prev is not None else "—",
                           "变化": chg, "判定": verdict})
    for c, cur in this_dist.items():
        prev_n = round(last_dist.get(c, 0) * 12) if last_dist else 0  # 上周约 12 条折算
        if last_dist and last_dist.get(c, 0) == 0 and dist.get(c, 0) >= 3:
            new_issues.append(c)

    # S4 治理排期
    top3 = [c for c, _ in sorted(this_dist.items(), key=lambda kv: -kv[1]) if dist.get(c, 0) > 0][:3]
    plan = [{"优先级": f"改进{i}", "根因": c, "本周占比": f"{this_dist[c]:.0%}",
             "治理动作": ACTIONS.get(c, ""), "责任人": "（周会指派）", "完成时点": "下周五"}
            for i, c in enumerate(top3, 1)]
    small_sample = len(negatives) < MIN_SAMPLE

    alarms = []
    if this_dist.get("质量问题(安全)", 0) > 0.10:
        alarms.append(f"质量安全类占 {this_dist['质量问题(安全)']:.0%}（>10%）→ 当日下架自查")
    if event_note and new_issues:
        alarms.append(f"新发根因 {('、'.join(new_issues))}，对照事件「{event_note}」排查因果")
    if malicious:
        alarms.append(f"筛查出 {len(malicious)} 条疑似恶意差评 → 不进归因，转 dispute-evidence-pack-flow 申诉")
    if small_sample:
        alarms.append(f"样本 {len(negatives)} 条 < {MIN_SAMPLE}，占比仅作参考，逐条处置为主")

    summary = {
        "店铺": shop, "平台": platform, "周期": period,
        "差评总数": len(negatives), "有效归因": len(detail),
        "加权总数": weight_sum, "疑似恶意": len(malicious),
        "首要根因": top3[0] if top3 else "—", "首要根因占比": f"{this_dist.get(top3[0], 0):.0%}" if top3 else "—",
        "趋势判定噪声带": NOISE_BAND, "样本下限": MIN_SAMPLE,
        "告警": "；".join(alarms) or "无",
        "事件备注": event_note or "（未提供）",
        "编排说明": "S1 筛查 → S2 归因(同源 negative-review-rootcause) → S3 趋势对比 → S4 治理排期；数值以脚本输出为准",
    }

    at.ensure_outdir(outdir)
    chart = at.bar_chart(
        os.path.join(outdir, "根因趋势.png"),
        [t["根因"] for t in trend_rows][::-1],
        [dist.get(t["根因"], 0) for t in trend_rows][::-1],
        title="本周根因加权分布（追评×2）", horizontal=True,
    )
    xlsx = at.write_excel(
        os.path.join(outdir, "差评归因治理表.xlsx"),
        {
            "趋势对比": trend_rows,
            "治理排期": plan or [{"优先级": "—", "根因": "（本周无有效差评）"}],
            "归因明细": detail,
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"趋势对比": {"判定": "contains:恶化"},
                    "归因明细": {"根因": "contains:安全"}},
        widths={"治理排期": {"治理动作": 44}, "归因明细": {"摘要": 32}},
    )
    js = at.write_json({"summary": summary, "trend": trend_rows, "plan": plan,
                        "malicious": malicious, "generated_at": at.stamp(),
                        "note": "机器归因与对比；因果解释由模型按 prompt.txt 复核"},
                       os.path.join(outdir, "rootcause_flow.json"))
    return {"files": [xlsx, chart, js], "summary": summary}


def main():
    ap = argparse.ArgumentParser(description="差评根因归因工作流")
    ap.add_argument("--input", help="输入 JSON（shop/period/negatives/last_week_dist）")
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
    s = r["summary"]
    print(f"{s['差评总数']} 条差评（有效 {s['有效归因']}，疑似恶意 {s['疑似恶意']}）：首要根因「{s['首要根因']}」占 {s['首要根因占比']}")
    print(f"告警：{s['告警']}")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
