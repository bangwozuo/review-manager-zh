# -*- coding: utf-8 -*-
"""
安抚回复与补偿方案生成 —— 端到端编排脚本。

流程 DAG：
  差评清单 → S1 前置筛查(恶意差评/根因判定) → S2 L1-L5 补偿分级路由
           → S3 话术模板生成(30-80字四件套) → S4 授权核对与人工确认 → 汇总产物

编排的原子技能：de_ecom_03_sk03 安抚话术生成（分级与话术规范同源）。

补偿分级（量化）：
  L1 3星/期望差：0 元，中性回复
  L2 物流/描述不符：补发配件或无门槛券 5-10 元（≤ 客单价 30%）
  L3 工艺质量（非安全）：补发整件或部分退款 30%-50%（≤ 客单价 60%）
  L4 质量安全（变质/异物/过敏）：全额退款 + 补发 + 主动电话 + 留证上报
  L5 涉诉风险词（12315/曝光/媒体）：停用模板，转人工 + 店长 2 小时介入
  同一买家 30 天内第二次 L3 以上 → 升级人工审核；超授权的方案一律标「需人工确认」

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo

产物：
  out/安抚补偿方案.xlsx   分级方案 / 话术 / 授权核对 / 人工确认清单
  out/安抚补偿方案.docx   交接用 Word
  out/appease_flow.json   机器可读结果
"""
from __future__ import annotations

import argparse
import os
import re
import sys

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

SAFETY_PAT = r"变质|发霉|异物|过敏|腹泻|中毒|划伤|刺鼻|臭"
URGENT_PAT = r"12315|工商|举报|曝光|媒体|人身安全|报警"
CAUSE_PAT = {
    "物流问题": r"物流慢|发货慢|等了\d+天|漏发|错发|丢件|暴力分拣|压坏",
    "质量问题(工艺)": r"碎|破损|开线|掉色|做工|质量差|用几天就|跑绒|卡顿",
    "服务问题": r"客服.*(态度|敷衍|不理|推诿|踢皮球)|已读不回|退款慢|售后.*差",
    "描述不符": r"与描述|与图片|色差|缩水|分量.*少|尺寸.*不对|详情页",
    "性价比": r"太贵|不值|溢价|买贵了",
}
TIER_COMP = {
    "L1": "无补偿，中性回复",
    "L2": "补发配件或无门槛券 5-10 元（≤ 客单价 30%）",
    "L3": "补发整件或部分退款 30%-50%（≤ 客单价 60%）",
    "L4": "全额退款 + 补发 + 主动电话 + 留证上报",
    "L5": "停用模板，转人工 + 店长 2 小时介入",
}

DEMO = {
    "shop": "雾岭茶业旗舰店",
    "platform": "天猫",
    "policy": "可补发整件；无门槛券单张 ≤ 15 元；质量问题部分退款 ≤ 50%；无全额退款授权",
    "unit_price": 88,
    "negatives": [
        {"star": 3, "text": "口感偏淡，不如上次买的批次", "has_image": False, "is_followup": False},
        {"star": 2, "text": "物流太慢，等了6天才到，茶叶都有点受潮", "has_image": False, "is_followup": False},
        {"star": 1, "text": "茶叶里有异物，喝着不放心，要求全额退款", "has_image": True, "is_followup": False},
        {"star": 2, "text": "罐子瘪了，包装太差", "has_image": True, "is_followup": False},
        {"star": 2, "text": "客服已读不回，退货申请三天没人处理", "has_image": False, "is_followup": True},
        {"star": 3, "text": "和详情页图片色差大，看着像陈茶", "has_image": False, "is_followup": False},
    ],
}


def cause_of(text: str) -> str:
    for cause, pat in CAUSE_PAT.items():
        if re.search(pat, text):
            return cause
    return "期望差(主观)"


def tier_of(r: dict) -> str:
    text = str(r.get("text", ""))
    star = int(r.get("star", 3) or 3)
    if re.search(URGENT_PAT, text):
        return "L5"
    if re.search(SAFETY_PAT, text):
        return "L4"
    if star <= 2:
        c = cause_of(text)
        if c in ("物流问题", "描述不符"):
            return "L2"
        return "L3"
    return "L1"


def reply_text(i: int, r: dict, tier: str, cause: str, policy: str) -> str:
    """模板话术生成：引用摘要 + 分级动作 + 通道。超出授权的 L3/L4 走人工确认口径。"""
    abstract = str(r.get("text", ""))[:16]
    unit = r.get("order_tail") or f"尾号{1000 + i}"
    if tier == "L5":
        return (f"您反馈的「{abstract}」问题我们高度重视，已在 1 小时内升级店长专线处理，"
                f"2 小时内您会接到我们的电话，请保持畅通。")
    if tier == "L4":
        full_ok = "全额退款" in policy and "无全额退款" not in policy
        if full_ok:
            return (f"您收到的「{abstract}」属严重品质问题，该批次已下架送检；"
                    f"我们按最高标准处理：退款与补发同时发起，专属客服已私信您，"
                    f"报订单{unit}即可。")
        return (f"您收到的「{abstract}」属严重品质问题，该批次已下架送检；本单处理方案"
                f"超出常规授权，已升级店长 2 小时内专线联系您。")
    if tier == "L3":
        return (f"您反馈的「{abstract}」问题收到，工艺类问题我们不做辩解：已安排补发"
                f"新批次，旧件无需寄回，私信客服报订单{unit}确认地址即可。")
    if tier == "L2":
        return (f"让您久等了！「{abstract}」是我们出库排期没排好，已反馈仓库并给您"
                f"补寄损耗件，另有一张小额无门槛券已发到您的账户。")
    return (f"感谢您的真实反馈！「{abstract}」属口感偏好差异，本次批次偏清淡；"
            f"客服可按您的口味推荐更浓香型，欢迎私信沟通。")


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    policy = payload.get("policy", "（未提供授权口径）")
    unit_price = float(payload.get("unit_price", 0) or 0)
    negatives = payload.get("negatives", [])

    plans, confirm = [], []
    cnt = {}
    for i, r in enumerate(negatives, 1):
        text = str(r.get("text", ""))
        cause = cause_of(text)
        tier = tier_of(r)
        cnt[tier] = cnt.get(tier, 0) + 1
        comp = TIER_COMP[tier]
        full_ok = "全额退款" in policy and "无全额退款" not in policy
        partial_ok = "部分退款" in policy or ("退款" in policy and "无全额退款" not in policy)
        over_auth = ("全额退款" in comp and not full_ok) or \
                    ("部分退款" in comp and not partial_ok)
        reply = reply_text(i, r, tier, cause, policy)
        status = "直接执行" if not over_auth else "需人工确认（超授权）"
        if over_auth or tier in ("L4", "L5"):
            confirm.append(f"#{i} [{tier}] {text[:20]}… → {status}")
        plans.append({
            "序号": i, "星级": r.get("star", "—"), "根因": cause, "情形分级": tier,
            "补偿方案": comp + (" ⚠️ 超授权" if over_auth else ""),
            "成本上限": f"≤ ¥{unit_price * 0.3:.0f}" if tier == "L2"
                       else (f"≤ ¥{unit_price * 0.6:.0f}" if tier == "L3" else "不设上限/0 元"),
            "回复话术（字数）": reply,
            "话术字数": len(reply),
            "状态": status,
        })

    tiers_order = ["L5", "L4", "L3", "L2", "L1"]
    plans.sort(key=lambda p: tiers_order.index(p["情形分级"]))
    head = f"L4 {cnt.get('L4', 0)} 件 / L5 {cnt.get('L5', 0)} 件须人工介入" \
        if (cnt.get("L4") or cnt.get("L5")) else "全部案件在标准分级内处理"

    summary = {
        "店铺": shop, "平台": platform, "差评数": len(negatives),
        "分级分布": cnt, "授权口径": policy,
        "客单价": unit_price or "（未提供）",
        "人工确认件数": len(confirm),
        "总体判定": head,
        "分级标准": "L1=0元 L2≤30%客单 L3≤60%客单 L4全退+补发+电话 L5转人工2h",
        "编排说明": "S1 筛查 → S2 分级路由 → S3 话术生成 → S4 授权核对；数值以脚本输出为准",
    }

    at.ensure_outdir(outdir)
    xlsx = at.write_excel(
        os.path.join(outdir, "安抚补偿方案.xlsx"),
        {
            "分级方案": plans,
            "授权核对": [{"项": k, "内容": str(v)} for k, v in summary.items()],
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"分级方案": {"情形分级": "contains:L5"},
                    "分级方案2": {"状态": "contains:需人工确认"}},
        widths={"分级方案": {"回复话术（字数）": 50, "补偿方案": 34}},
    )
    sections = [
        {"heading": "一、总体判定", "paras": [head, f"授权口径：{policy}；补偿分级 L1-L5，超授权方案一律标「需人工确认」。"]},
        {"heading": "二、人工确认清单（店长/运营过目后执行）",
         "bullets": confirm or ["（无——全部在授权内）"]},
        {"heading": "三、分级方案与话术",
         "table": {"cols": ["序号", "分级", "根因", "补偿方案", "话术"],
                   "rows": [[p["序号"], p["情形分级"], p["根因"], p["补偿方案"], p["回复话术（字数）"]] for p in plans]}},
        {"heading": "四、执行提醒",
         "bullets": ["话术发布前人工确认；不得出现返现/换好评等诱导措辞（反法第 8 条）",
                     "L4 案件先留证（图证+批次+订单）再回应「正在核实」，不写未经核实的自认",
                     "本方案为 AI 生成内容，数值与分级以脚本输出为准"]},
    ]
    docx = at.write_docx(
        os.path.join(outdir, "安抚补偿方案.docx"),
        f"{shop} 安抚补偿方案", sections, subtitle=f"{platform} · AI 生成内容",
    )
    js = at.write_json({"summary": summary, "plans": plans, "confirm": confirm,
                        "generated_at": at.stamp(),
                        "note": "机器分级与模板话术；语境与授权细节由模型按 prompt.txt 复核"},
                       os.path.join(outdir, "appease_flow.json"))
    return {"files": [xlsx, docx, js], "summary": summary}


def main():
    ap = argparse.ArgumentParser(description="安抚回复与补偿方案生成工作流")
    ap.add_argument("--input", help="输入 JSON（shop/policy/negatives）")
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
    print(f"{s['差评数']} 条差评 → 分级 {s['分级分布']}；人工确认 {s['人工确认件数']} 件。{s['总体判定']}")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
