# -*- coding: utf-8 -*-
"""
差评根因分类 —— 词表规则树归因器（六类根因 + 加权分布 + 下周改进项）。

职责边界：本脚本只做**负面词表命中、规则树归因、加权统计与产物生成**（机器的强项）。
根因间的真实因果链（如「物流慢导致坏果」的连带归因）、与 SKU 批次的交叉验证
由模型按 prompt.txt 完成（这是模型的强项）。

归因规则（量化标准）：
  六类根因：质量问题 / 物流问题 / 服务问题 / 描述不符 / 性价比 / 期望差（主观）
  规则树：先判质量安全（变质/异物/过敏，最高优先级），再物流，再服务，
          再描述不符（与详情页承诺冲突），再性价比；都不命中 → 期望差
  权重：追评比首评更可信，追评计入权重 ×2；带图差评另计「图证数」
  治理阈值：单一根因占比 > 40% → 列为首要治理专项；质量安全类占比 > 10% → 当日立即处理

用法：
  python rootcause.py --input input.json --outdir out
  python rootcause.py --demo

产物：
  out/差评根因分类.xlsx   根因明细 / 分布统计 / 下周改进项
  out/根因分布.png        六类根因占比柱状图
  out/rootcause.json      机器可读结果（供工作流读取）
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from collections import Counter

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

FOLLOWUP_WEIGHT = 2      # 追评权重 ×2（更可信）
TOP_CAUSE专项_THRESHOLD = 0.40  # 单一根因占比 > 40% → 首要治理专项
SAFETY_ALARM = 0.10      # 质量安全类占比 > 10% → 当日立即处理

# 词表：根因 → 负面词。按规则树顺序匹配（先安全，后物流/服务/描述/性价比，兜底期望差）。
CAUSE_WORDS = {
    "质量问题(安全)": r"变质|发霉|过期|异物|虫|头发|腹泻|过敏|中毒|异味|臭",
    "质量问题(工艺)": r"碎|破损|漏液|掉色|开线|做工粗糙|以次充好|假货|仿品|用几天就坏|质量差",
    "物流问题": r"物流慢|发货慢|迟迟不发货|等了\d+天|催了|暴力分拣|压坏|挤压|漏发|错发|丢件|迟迟不到|包装.*破",
    "服务问题": r"客服.*(态度|敷衍|不理|踢皮球|推诿)|态度差|爱答不理|已读不回|退款慢|售后.*差|无人处理|互相推",
    "描述不符": r"与描述|与图片|详情页|实物.*不符|色差|缩水|克重不足|分量.*少|尺寸.*不对|规格.*不对",
    "性价比": r"太贵|不值|溢价|买贵了|降价.*快|这个价.*不如|性价比",
}
CAUSE_ORDER = list(CAUSE_WORDS.keys())
FALLBACK = "期望差(主观)"
IMPROVE_ACTIONS = {
    "质量问题(安全)": "当日下架留证批次并自查，2 小时内响应全部涉事订单；对接品控出整改报告",
    "质量问题(工艺)": "本周内出供应商整改单，破损率目标降至 0.5% 以下；包装加缓冲",
    "物流问题": "核对近 7 天发货时长分位数（P90 ≤ 48 小时），超时仓库排班整改；易碎品换包装",
    "服务问题": "客服质检抽查本周差评会话，退款响应时限收紧至 2 小时；补售后话术",
    "描述不符": "对照详情页逐项修正主图/参数/分量描述，色差类补实拍图",
    "性价比": "评估赠品或规格分层，不做直接降价；详情页补同克重对比口径",
    FALLBACK: "沉淀期望管理：详情页增加「预期说明」，回复中客观呈现口径",
}

DEMO = {
    "shop": "果真甜生鲜旗舰店",
    "platform": "拼多多",
    "period": "2026-09-22 ~ 2026-09-28",
    "negatives": [
        {"star": 1, "text": "芒果烂了三个，箱子里一股酒味，都变质了", "has_image": True, "is_followup": False},
        {"star": 1, "text": "吃了一个晚上拉肚子，里面有异物，必须投诉", "has_image": True, "is_followup": False},
        {"star": 2, "text": "暴力分拣压坏一半，包装太差", "has_image": True, "is_followup": False},
        {"star": 2, "text": "发货慢，等了5天还没到，催了两次", "has_image": False, "is_followup": False},
        {"star": 2, "text": "漏发一箱，客服已读不回，售后太差", "has_image": False, "is_followup": False},
        {"star": 3, "text": "实物和图片色差大，详情页说 5 斤实际分量不足", "has_image": False, "is_followup": False},
        {"star": 2, "text": "太贵了，这个价不如楼下超市", "has_image": False, "is_followup": False},
        {"star": 1, "text": "烂果率太高，第二箱还是坏的", "has_image": False, "is_followup": True},
        {"star": 3, "text": "物流太慢，到手都不新鲜了", "has_image": False, "is_followup": True},
        {"star": 2, "text": "客服推诿不处理烂果赔付", "has_image": False, "is_followup": False},
        {"star": 3, "text": "个头比描述小一圈", "has_image": False, "is_followup": False},
        {"star": 2, "text": "破损，垫纸就两层", "has_image": True, "is_followup": False},
    ],
}


def attribute(text: str):
    """规则树归因：按 CAUSE_ORDER 顺序命中即返回（根因, 次要标签列表）。"""
    primary, secondary = None, []
    for cause in CAUSE_ORDER:
        if re.search(CAUSE_WORDS[cause], text):
            if primary is None:
                primary = cause
            else:
                secondary.append(cause)
    if primary is None:
        primary = FALLBACK
    return primary, secondary


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    period = payload.get("period", "")
    negatives = payload.get("negatives", [])

    detail, dist = [], Counter()
    weight_sum = 0
    for i, r in enumerate(negatives, 1):
        text = str(r.get("text", ""))
        primary, secondary = attribute(text)
        weight = FOLLOWUP_WEIGHT if r.get("is_followup") else 1
        weight_sum += weight
        dist[primary] += weight
        detail.append({
            "序号": i,
            "星级": r.get("star", "—"),
            "根因(主)": primary,
            "根因(次)": "、".join(secondary) or "—",
            "权重": f"×{weight}" if weight > 1 else "×1",
            "图证": "有" if r.get("has_image") else "无",
            "评价摘要": text[:30] + ("…" if len(text) > 30 else ""),
        })

    total_w = weight_sum or 1
    dist_rows = [{"根因": c, "加权条数": dist.get(c, 0),
                  "占比": f"{dist.get(c, 0) / total_w:.0%}",
                  "治理动作": IMPROVE_ACTIONS.get(c, "")}
                 for c in CAUSE_WORDS] + \
                [{"根因": FALLBACK, "加权条数": dist.get(FALLBACK, 0),
                  "占比": f"{dist.get(FALLBACK, 0) / total_w:.0%}",
                  "治理动作": IMPROVE_ACTIONS[FALLBACK]}]
    dist_rows.sort(key=lambda x: -x["加权条数"])

    top = dist_rows[0]
    top_ratio = top["加权条数"] / total_w
    safety_ratio = dist.get("质量问题(安全)", 0) / total_w
    alarms = []
    if top_ratio > TOP_CAUSE专项_THRESHOLD:
        alarms.append(f"「{top['根因']}」占 {top_ratio:.0%}（> {TOP_CAUSE专项_THRESHOLD:.0%} 阈值）→ 列为首要治理专项")
    if safety_ratio > SAFETY_ALARM:
        alarms.append(f"质量安全类占 {safety_ratio:.0%}（> {SAFETY_ALARM:.0%}）→ 当日下架留证批次并自查，勿拖到周会")
    if not alarms:
        alarms.append("分布均衡：按 Top3 根因排入下周改进，每项带数字目标")

    summary = {
        "店铺": shop, "平台": platform, "统计周期": period,
        "差评总数": len(negatives), "加权总数": weight_sum,
        "首要根因": top["根因"], "首要根因占比": top["占比"],
        "质量安全占比": f"{safety_ratio:.0%}",
        "治理判定": "；".join(alarms),
        "说明": "词表规则树机器归因；连带因果与 SKU 批次交叉验证由模型按 prompt.txt 复核；数值以脚本输出为准，不要自己算",
    }

    at.ensure_outdir(outdir)
    xlsx = at.write_excel(
        os.path.join(outdir, "差评根因分类.xlsx"),
        {
            "分布统计": dist_rows,
            "根因明细": detail,
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"分布统计": {"根因": "contains:安全"},
                    "根因明细": {"根因(主)": "contains:安全"}},
        widths={"分布统计": {"治理动作": 46}, "根因明细": {"评价摘要": 34}},
    )
    nonzero = [r for r in dist_rows if r["加权条数"] > 0]
    chart = at.bar_chart(
        os.path.join(outdir, "根因分布.png"),
        [r["根因"] for r in nonzero][::-1],
        [r["加权条数"] for r in nonzero][::-1],
        title="差评根因加权分布（追评×2）", horizontal=True,
    )
    js = at.write_json({"summary": summary, "distribution": dist_rows, "detail": detail,
                        "generated_at": at.stamp(),
                        "note": "机器归因结果；因果链与批次交叉由模型按 prompt.txt 复核"},
                       os.path.join(outdir, "rootcause.json"))
    return {"files": [xlsx, chart, js], "summary": summary}


def main():
    ap = argparse.ArgumentParser(description="差评根因分类 —— 规则树归因")
    ap.add_argument("--input", help="输入 JSON（shop/platform/period/negatives）")
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
    print(f"{s['差评总数']} 条差评：首要根因「{s['首要根因']}」占 {s['首要根因占比']}，质量安全类 {s['质量安全占比']}")
    print(f"治理判定：{s['治理判定']}")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
