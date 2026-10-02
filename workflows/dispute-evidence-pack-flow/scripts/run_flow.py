# -*- coding: utf-8 -*-
"""
纠纷证据包整理 —— 恶意差评申诉端到端编排脚本。

流程 DAG：
  案件输入 → S1 恶意特征检测 → S2 证据完整度打分(五项证据)
           → S3 可提交判定(≥80% 提交 / 60-80% 补强 / <60% 不提交)
           → S4 申诉信要素表 + 补证清单 → 汇总产物

申诉规则（量化，来源见 skills/platform-rules-qa 内置规则库 R01）：
  申诉成功率行业经验约 40-60%，差异全在证据完整度
  五项证据：订单记录(20分) 完整聊天记录含时间戳(30分) 物流凭证(20分)
            评价原文截图含账号与时间(20分) 买家行为佐证(10分)
  完整度 = 实得分 / 100；< 60% 严禁提交（驳回且占用当日申诉次数）

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo

产物：
  out/申诉证据包.xlsx   恶意特征检测 / 证据打分 / 申诉信要素 / 汇总
  out/申诉证据包.docx   交接用 Word（可提交判定与补证清单）
  out/dispute_flow.json 机器可读结果
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

SUBMIT_LINE = 0.80      # ≥ 80% 可提交
REINFORCE_LINE = 0.60   # 60-80% 补强后提交；< 60% 不提交

# 五项证据：名称 / 分值 / 核对要点
EVIDENCE_ITEMS = [
    ("订单记录", 20, "证明真实交易或证明与评价者无关联交易"),
    ("完整聊天记录", 30, "含时间戳的原始截图，不裁剪；敲诈勒索型申诉的核心证据"),
    ("物流凭证", 20, "签收/拒收/未发货面单，「未收货差评」型申诉的核心证据"),
    ("评价原文截图", 20, "含发布时间与账号信息，用于平台定位违规评价"),
    ("买家行为佐证", 10, "历史差评记录/拒收记录等惯发性证据（如有）"),
]

# 恶意特征检测：特征名 / 正则 / 申诉类型
MALICIOUS_SIGNS = [
    ("索要返现删评", r"加.{0,3}(微信|V|v)|转账|红包.{0,6}(删|改)|返.{0,4}元.{0,4}(删|改)|撤差评", "敲诈勒索型"),
    ("同行攻击", r"低价|同款.{0,6}便宜|比.{0,6}店|同行", "同行攻击型"),
    ("辱骂人身攻击", r"妈的|全家|骗子.{0,4}死|垃圾店", "辱骂型"),
    ("无交易差评", r"没买过|没有订单", "无交易型"),
]

DEMO = {
    "shop": "临风运动户外店",
    "platform": "抖音",
    "cases": [
        {
            "case_id": "AP-20260929-01",
            "review_text": "垃圾店，东西根本没用过就给差评，加V给你转50删了",
            "has_order": True, "chat_screenshot": True, "logistics_proof": False,
            "review_screenshot": True, "behavior_proof": True,
        },
        {
            "case_id": "AP-20260929-02",
            "review_text": "同款比你们便宜一半，大家别来这家",
            "has_order": True, "chat_screenshot": False, "logistics_proof": True,
            "review_screenshot": True, "behavior_proof": False,
        },
        {
            "case_id": "AP-20260929-03",
            "review_text": "尺码偏小，客服换了很顺利，就是等的有点久",
            "has_order": True, "chat_screenshot": True, "logistics_proof": True,
            "review_screenshot": True, "behavior_proof": False,
        },
    ],
}


def malicious_signs(text: str):
    signs = []
    for name, pat, case_type in MALICIOUS_SIGNS:
        if re.search(pat, text):
            signs.append({"特征": name, "申诉类型": case_type})
    return signs


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    cases = payload.get("cases", [])

    rows, letters = [], []
    for c in cases:
        cid = c.get("case_id", "—")
        text = str(c.get("review_text", ""))
        signs = malicious_signs(text)
        provided = {
            "订单记录": bool(c.get("has_order")),
            "完整聊天记录": bool(c.get("chat_screenshot")),
            "物流凭证": bool(c.get("logistics_proof")),
            "评价原文截图": bool(c.get("review_screenshot")),
            "买家行为佐证": bool(c.get("behavior_proof")),
        }
        got = sum(w for (name, w, _), ok in zip(EVIDENCE_ITEMS, provided.values()) if ok)
        total = sum(w for _, w, _ in EVIDENCE_ITEMS)
        ratio = got / total
        missing = [name for (name, _, _), ok in zip(EVIDENCE_ITEMS, provided.values()) if not ok]
        if not signs:
            verdict = "→ 转安抚流（未检出恶意特征，属真实差评）"
            missing = "—"
        elif ratio >= SUBMIT_LINE:
            verdict = "✅ 可提交"
            missing = "、".join(missing) or "—"
        elif ratio >= REINFORCE_LINE:
            verdict = "🟡 补强后提交"
            missing = "、".join(missing) or "—"
        else:
            verdict = "⛔ 不提交（证据不全）"
            missing = "、".join(missing) or "—"
        rows.append({
            "案件ID": cid,
            "恶意特征": "、".join(s["特征"] for s in signs) or "未检出",
            "申诉类型": signs[0]["申诉类型"] if signs else ("误判澄清（差评描述与订单可对上）" if not signs else "—"),
            "证据得分": f"{got}/{total}",
            "完整度": f"{ratio:.0%}",
            "判定": verdict,
            "缺项": "、".join(missing) or "—",
        })
        if signs and ratio >= REINFORCE_LINE:
            letters.append({
                "案件ID": cid, "申诉类型": signs[0]["申诉类型"], "完整度": f"{ratio:.0%}",
                "申诉信要素": (f"①事实陈述：买家评价「{text[:20]}…」与本店订单记录无关联交易"
                              f"（{'有' if provided['订单记录'] else '无'}订单比对结果）；"
                              f"②证据清单：{'；'.join(n for n, ok in provided.items() if ok)}；"
                              f"③诉求：判定恶意评价并删除展示；④承诺：以上材料真实，愿承担虚假申诉责任。"),
            })

    n_submit = sum(1 for r in rows if r["判定"] == "✅ 可提交")
    n_hold = sum(1 for r in rows if "补强" in r["判定"])
    n_stop = sum(1 for r in rows if "不提交" in r["判定"])
    n_normal = sum(1 for r in rows if r["判定"].startswith("→"))

    summary = {
        "店铺": shop, "平台": platform, "案件数": len(cases),
        "可提交": n_submit, "补强后提交": n_hold, "不提交": n_stop,
        "转安抚流": n_normal,
        "提交线": SUBMIT_LINE, "补强线": REINFORCE_LINE,
        "成功率基准": "行业经验约 40-60%，差异全在证据完整度",
        "提醒": "驳回两次以上暂停提交，人工复核案件定性；证据不全的提交会占用当日申诉次数",
        "编排说明": ("S1 恶意特征检测 → S2 证据五项打分 → S3 可提交判定 → S4 申诉信要素；"
                     "条款依据见 platform-rules-qa 规则库 R01；数值以脚本输出为准"),
    }

    at.ensure_outdir(outdir)
    xlsx = at.write_excel(
        os.path.join(outdir, "申诉证据包.xlsx"),
        {
            "案件判定": rows,
            "申诉信要素": letters or [{"案件ID": "—", "申诉类型": "—", "完整度": "—", "申诉信要素": "（无可提交案件）"}],
            "证据核对表": [{"证据项": n, "分值": w, "核对要点": p} for n, w, p in EVIDENCE_ITEMS],
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"案件判定": {"判定": "contains:不提交"}},
        widths={"案件判定": {"缺项": 24}, "申诉信要素": {"申诉信要素": 60}},
    )
    sections = [
        {"heading": "一、总体判定",
         "paras": [f"案件 {len(cases)} 件：可提交 {n_submit}、补强后提交 {n_hold}、不提交 {n_stop}、转安抚流 {n_normal}。",
                    f"提交线 {SUBMIT_LINE:.0%}、补强线 {REINFORCE_LINE:.0%}；申诉成功率行业经验约 40-60%，差异全在证据完整度。"]},
        {"heading": "二、逐案判定",
         "table": {"cols": ["案件ID", "恶意特征", "证据得分", "完整度", "判定", "缺项"],
                   "rows": [[r["案件ID"], r["恶意特征"], r["证据得分"], r["完整度"], r["判定"], r["缺项"]] for r in rows]}},
        {"heading": "三、申诉信要素（模板，提交前人工补充订单号并核对）",
         "bullets": [f"{l['案件ID']}（{l['申诉类型']}，完整度 {l['完整度']}）" for l in letters] or ["（无可提交案件）"]},
        {"heading": "四、执行提醒",
         "bullets": ["证据不全不提交——驳回会占用当日申诉次数；驳回两次以上暂停并人工复核",
                     "聊天记录须原始截图含时间戳，不裁剪；提交前脱敏第三方隐私信息",
                     "本判定为 AI 生成内容，申诉结果以平台审核为准"]},
    ]
    docx = at.write_docx(
        os.path.join(outdir, "申诉证据包.docx"),
        f"{shop} 恶意差评申诉证据包", sections, subtitle=f"{platform} · AI 生成内容",
    )
    js = at.write_json({"summary": summary, "cases": rows, "letters": letters,
                        "generated_at": at.stamp(),
                        "note": "机器打分与判定；申诉信撰写由模型按 prompt.txt 细化"},
                       os.path.join(outdir, "dispute_flow.json"))
    return {"files": [xlsx, docx, js], "summary": summary}


def main():
    ap = argparse.ArgumentParser(description="纠纷证据包整理工作流")
    ap.add_argument("--input", help="输入 JSON（shop/platform/cases）")
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
    print(f"{s['案件数']} 件案件：可提交 {s['可提交']} / 补强 {s['补强后提交']} / 不提交 {s['不提交']}")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
