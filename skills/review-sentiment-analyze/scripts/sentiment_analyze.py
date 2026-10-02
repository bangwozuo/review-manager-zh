# -*- coding: utf-8 -*-
"""
评价情感分析 —— 星级 + 词表五级情感判定器（带图差评优先、追评加权、四维标签词频）。

职责边界：本脚本只做**星级映射、负面词表命中、四维标签词频统计与产物生成**（机器的强项）。
反讽/阴阳怪气语境（如「真是太好了呢，三天就碎了」）、刷出来的假好评识别
由模型按 prompt.txt 复核（这是模型的强项）。

判定规则（量化标准）：
  情感五级映射：星 1-2 → 强负；星 3 → 负面；无文字且无星级 → 中性；星 4 → 正面；
                星 5 且（带图或文字 ≥ 15 字）→ 强正，否则计入正面
  处置优先级：带图差评（星 ≤ 3 且带图）= P0（影响转化最大，2 小时内响应）；
              强负 = P1（12 小时内）；负面 = P2（24 小时内）；追评差评自动升一级
  标签权重：追评比首评更可信，词频统计时追评命中权重 ×2
  健康度基准：好评率（4-5 星占比）≥ 95%；差评率（1-2 星占比）> 3% 须专项治理

用法：
  python sentiment_analyze.py --input input.json --outdir out
  python sentiment_analyze.py --demo

产物：
  out/评价情感分析.xlsx   分级明细 / 四维标签词频 / 汇总
  out/情感分布.png        五级情感占比饼图
  out/负面标签Top.png     负面标签频次柱状图
  out/sentiment.json      机器可读结果（供工作流读取）
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

GOOD_RATE_FLOOR = 0.95   # 好评率（4-5 星占比）下限
BAD_RATE_CEIL = 0.03     # 差评率（1-2 星占比）上限，超过须专项治理
FOLLOWUP_WEIGHT = 2      # 追评权重（比首评更可信）
STRONG_POS_MIN_CHARS = 15  # 5 星无图时，文字 ≥ 15 字才算强正
SLA = {"P0": "2 小时内响应", "P1": "12 小时内响应", "P2": "24 小时内响应"}

# 四维标签词表：维度 → (负面词, 正面词)。命中即打标，词频用于归因。
DIM_WORDS = {
    "口味/质量": {
        "neg": r"难吃|难喝|变质|发霉|过期|异味|口感差|口感差|碎|破损|漏液|质量差|做工粗糙|假货|仿品|以次充好|掉色|开线|用几天就|第二次买.*坏|可惜",
        "pos": r"好吃|好喝|新鲜|口感好|味道正|质量好|做工精细|真材实料|耐用|结实|正品|品质",
    },
    "服务": {
        "neg": r"客服.*(态度|敷衍|不理人|踢皮球)|态度差|爱答不理|敷衍|推诿|退款慢|售后.*差|无人处理|互相推|已读不回",
        "pos": r"客服.*(耐心|热情|及时)|态度好|秒回|售后.*到位|处理得快|退款快|贴心",
    },
    "物流": {
        "neg": r"物流慢|发货慢|迟迟不发货|等了.*天|催了.*次|暴力分拣|包装.*破|压坏|漏发|错发|丢件|迟迟不到",
        "pos": r"发货快|物流快|次日达|隔天到|包装完好|包装严实|准时送达",
    },
    "性价比": {
        "neg": r"太贵|不值|溢价|降价.*快|买贵了|分量.*少|缩水|克重不足|这个价.*不如",
        "pos": r"实惠|划算|性价比高|分量足|物超所值|便宜大碗|活动价入手",
    },
}

URGENT_WORDS = r"投诉|12315|工商|举报|退货退款|仅退款|骗子|假货|曝光|给差评|售后无门|腹泻|过敏|受伤|安全问题"

DEMO = {
    "shop": "川味鲜食品旗舰店",
    "platform": "天猫",
    "period": "2026-09-22 ~ 2026-09-28",
    "reviews": [
        {"star": 1, "text": "吃到变质发霉的，客服已读不回，再不处理就打12315投诉", "has_image": True, "is_followup": False},
        {"star": 2, "text": "包装压坏了半盒，物流太慢等了6天，客服态度还敷衍", "has_image": True, "is_followup": False},
        {"star": 1, "text": "难吃，口感差，感觉是假货，要求退货退款", "has_image": False, "is_followup": False},
        {"star": 3, "text": "味道一般，分量比上次少，这个价格不太值", "has_image": False, "is_followup": False},
        {"star": 2, "text": "漏发了两包，找客服三天才补发，太慢了", "has_image": False, "is_followup": False},
        {"star": 3, "text": "物流慢，等了一个星期才到", "has_image": False, "is_followup": True},
        {"star": 1, "text": "第二次买了还是碎的，做工越来越差", "has_image": False, "is_followup": True},
        {"star": 4, "text": "味道不错，发货也挺快，就是外包装有点压痕", "has_image": False, "is_followup": False},
        {"star": 5, "text": "麻辣味很正，越嚼越香，包装严实没漏，回购第三次了", "has_image": True, "is_followup": False},
        {"star": 5, "text": "新鲜好吃，客服很耐心地推荐了口味，物流次日达", "has_image": False, "is_followup": False},
        {"star": 5, "text": "性价比高，分量足，活动价入手太划算了", "has_image": False, "is_followup": False},
        {"star": 4, "text": "口感还行，包装完好，就是有点小贵", "has_image": False, "is_followup": False},
        {"star": 5, "text": "好吃", "has_image": False, "is_followup": False},
        {"star": 4, "text": "质量可以，物流稍慢等了4天", "has_image": False, "is_followup": False},
        {"star": 2, "text": "客服推诿不给退款，售后太差", "has_image": False, "is_followup": True},
        {"star": 5, "text": "回购N次的老店，品质一直在线，发货快", "has_image": True, "is_followup": False},
        {"star": 3, "text": "一般", "has_image": False, "is_followup": False},
        {"star": 5, "text": "味道正，包装好，送人也体面", "has_image": False, "is_followup": False},
        {"star": 1, "text": "过敏了，肚子不舒服，安全问题必须曝光", "has_image": False, "is_followup": False},
        {"star": 4, "text": "味道好，物流快", "has_image": False, "is_followup": False},
    ],
}


def classify_level(review: dict) -> str:
    """情感五级映射。中=无文字（且无有效星级信息时兜底中性）。"""
    star = review.get("star")
    text = str(review.get("text", "")).strip()
    if not text and star is None:
        return "中"
    star = 3 if star is None else int(star)
    if star <= 2:
        return "强负"
    if star == 3:
        return "负面"
    if star == 5:
        if review.get("has_image") or len(text) >= STRONG_POS_MIN_CHARS:
            return "强正"
        return "正面"
    return "正面"  # star == 4


def priority(level: str, review: dict) -> str:
    """处置优先级：带图差评 P0 > 强负 P1 > 负面 P2；追评差评自动升级。"""
    if level in ("强负", "负面") and review.get("has_image"):
        return "P0"
    if level == "强负":
        p = "P1"
    elif level == "负面":
        p = "P2"
    else:
        return "—"
    if review.get("is_followup") and p != "P0":
        p = {"P1": "P0", "P2": "P1"}[p]
    return p


def tags_of(text: str) -> list:
    hits = []
    for dim, words in DIM_WORDS.items():
        if re.search(words["neg"], text) or re.search(words["pos"], text):
            hits.append(dim)
    return hits


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    period = payload.get("period", "")
    reviews = payload.get("reviews", [])

    detail, tag_freq = [], Counter()
    cnt = Counter()
    n = len(reviews) or 1
    for i, r in enumerate(reviews, 1):
        text = str(r.get("text", ""))
        level = classify_level(r)
        pri = priority(level, r)
        tags = tags_of(text)
        weight = FOLLOWUP_WEIGHT if r.get("is_followup") else 1
        urgent = bool(re.search(URGENT_WORDS, text))
        for t in tags:
            if level in ("强负", "负面"):
                tag_freq[t] += weight  # 只有负面评价参与负面归因
        cnt[level] += 1
        detail.append({
            "序号": i,
            "星级": r.get("star", "—"),
            "情感分级": level,
            "处置优先级": pri,
            "响应时限": SLA.get(pri, "—"),
            "追评": "是（权重×2）" if r.get("is_followup") else "否",
            "带图": "是" if r.get("has_image") else "否",
            "四维标签": "、".join(tags) or "—",
            "涉诉风险词": "⚠️ 命中" if urgent else "",
            "评价摘要": text[:30] + ("…" if len(text) > 30 else ""),
        })

    good = cnt["强正"] + cnt["正面"]
    bad = cnt["强负"]
    total = len(reviews)
    good_rate = good / n
    bad_rate = bad / n
    health = []
    health.append(("好评率", f"{good_rate:.1%}", "达标" if good_rate >= GOOD_RATE_FLOOR
                   else f"低于基准 {GOOD_RATE_FLOOR:.0%}，须提升"))
    health.append(("差评率(1-2星)", f"{bad_rate:.1%}", "达标" if bad_rate <= BAD_RATE_CEIL
                   else f"超过 {BAD_RATE_CEIL:.0%}，须专项治理"))
    health.append(("带图差评", f"{sum(1 for d in detail if d['带图'] == '是' and d['星级'] not in (4, 5, '—'))} 条",
                   "P0 优先处理，对转化影响最大"))
    health.append(("追评差评", f"{sum(1 for r in reviews if r.get('is_followup') and r.get('star', 5) <= 3)} 条",
                   "追评比首评权重更高，已按 ×2 计入归因"))
    n_urgent = sum(1 for d in detail if d["涉诉风险词"])
    health.append(("涉诉风险评价", f"{n_urgent} 条", "含 12315/投诉/安全词，转人工即时处理"))

    top_tags = tag_freq.most_common(8)
    if top_tags and tag_freq[top_tags[0][0]] / max(1, sum(tag_freq.values())) > 0.4:
        focus = f"首要治理项：{top_tags[0][0]}（占负面标签 {tag_freq[top_tags[0][0]] / sum(tag_freq.values()):.0%}，> 40% 阈值）"
    else:
        focus = "负面标签分散，按 Top3 逐项排期治理"

    summary = {
        "店铺": shop, "平台": platform, "统计周期": period,
        "评价总数": total,
        "情感分布": {k: cnt.get(k, 0) for k in ("强负", "负面", "中", "正面", "强正")},
        "好评率": f"{good_rate:.1%}", "差评率": f"{bad_rate:.1%}",
        "健康度判定": focus,
        "说明": "星级+词表机器判定；反讽/水军语境由模型按 prompt.txt 复核；数值以脚本输出为准，不要自己算",
    }

    at.ensure_outdir(outdir)
    xlsx = at.write_excel(
        os.path.join(outdir, "评价情感分析.xlsx"),
        {
            "分级明细": detail,
            "四维标签词频": [{"维度": k, "加权频次": v,
                              "占比": f"{v / max(1, sum(tag_freq.values())):.0%}"}
                             for k, v in top_tags] or [{"维度": "（无负面标签）", "加权频次": 0, "占比": ""}],
            "健康度对照": [{"指标": k, "实测": v, "判定": d} for k, v, d in health],
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"分级明细": {"处置优先级": "contains:P0"},
                    "健康度对照": {"判定": "contains:治理"}},
        widths={"分级明细": {"评价摘要": 34, "响应时限": 14}},
    )
    pie = at.pie_chart(
        os.path.join(outdir, "情感分布.png"),
        [f"{k}" for k in ("强负", "负面", "中", "正面", "强正") if cnt.get(k)],
        [cnt[k] for k in ("强负", "负面", "中", "正面", "强正") if cnt.get(k)],
        title=f"情感五级分布（n={total}）",
    )
    bar = at.bar_chart(
        os.path.join(outdir, "负面标签Top.png"),
        [k for k, _ in top_tags][::-1], [v for _, v in top_tags][::-1],
        title="负面四维标签加权词频（追评×2）", horizontal=True,
    )
    js = at.write_json({"summary": summary, "detail": detail, "tag_freq": dict(top_tags),
                        "generated_at": at.stamp(),
                        "note": "机器判定结果；反讽/刷评识别由模型按 prompt.txt 复核"},
                       os.path.join(outdir, "sentiment.json"))
    return {"files": [xlsx, pie, bar, js], "summary": summary,
            "p0_count": sum(1 for d in detail if d["处置优先级"] == "P0")}


def main():
    ap = argparse.ArgumentParser(description="评价情感分析 —— 五级情感 + 四维标签")
    ap.add_argument("--input", help="输入 JSON（shop/platform/period/reviews）")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--demo", action="store_true", help="用内置样例跑一遍")
    a = ap.parse_args()

    if a.demo:
        payload = DEMO
    elif a.input:
        payload = at.read_json(a.input)
    else:
        ap.error("需要 --input / --demo 之一")

    r = build(payload, a.outdir)
    s = r["summary"]
    print(f"{s['评价总数']} 条评价：好评率 {s['好评率']} / 差评率 {s['差评率']}，P0 带图差评 {r['p0_count']} 条")
    print(f"健康度：{s['健康度判定']}")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
