# -*- coding: utf-8 -*-
"""
评价流实时监控与分级 —— 端到端编排脚本。

流程 DAG：
  数据接入 → S1 完整性校验 → S2 五级情感分级(review-sentiment-analyze 规则)
           → S3 告警分级与 SLA 排期 → S4 处置分发(安抚/申诉/人工) → 汇总产物

编排的原子技能：de_ecom_03_sk01 评价情感分析（分级规则同源）。

SLA 规则（量化）：
  P0 带图差评：2 小时内响应（转化影响最大）
  P1 强负无图：12 小时内响应；含涉诉风险词（12315/曝光/过敏）直接转人工
  P2 负面(3星)：24 小时内响应
  追评差评自动升一级；好评也要回复（30-80 字，24 小时内）
  今日未清零 P0 → 触发人工群内提醒

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo

产物：
  out/评价监控处置清单.xlsx   今日必办（按 P0→P2 排序）/ 好评回复 / 汇总
  out/评价监控处置清单.docx   交接用 Word（含 SLA 与转人工名单）
  out/monitor_flow.json       机器可读结果
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

SLA_HOURS = {"P0": 2, "P1": 12, "P2": 24, "好评": 24}
URGENT_WORDS = r"12315|工商|举报|曝光|媒体|骗子|假货|过敏|腹泻|中毒|受伤|人身安全"

DEMO = {
    "shop": "晨光母婴旗舰店",
    "platform": "天猫",
    "monitor_date": "2026-09-29",
    "reviews": [
        {"star": 1, "text": "奶瓶到手就裂了，还划到手，这种质量问题必须曝光", "has_image": True, "is_followup": False},
        {"star": 2, "text": "奶粉勺子没放，找客服三天没人理", "has_image": False, "is_followup": False},
        {"star": 1, "text": "宝宝用了红屁股，你们说食品级硅胶是骗人的吧", "has_image": True, "is_followup": True},
        {"star": 3, "text": "奶嘴流速太快，宝宝呛了好几次", "has_image": False, "is_followup": False},
        {"star": 2, "text": "发货太慢，等了5天，双11都比你快", "has_image": False, "is_followup": False},
        {"star": 4, "text": "质量不错，物流略慢", "has_image": False, "is_followup": False},
        {"star": 5, "text": "第二瓶回购，防胀气设计真的好用，客服还教了排气手法", "has_image": True, "is_followup": False},
        {"star": 3, "text": "包装一般，盒子有点瘪", "has_image": False, "is_followup": True},
    ],
}


def classify(r: dict):
    """与 skills/review-sentiment-analyze 同源的分级规则。"""
    star = r.get("star")
    text = str(r.get("text", ""))
    if not text and star is None:
        level = "中"
    elif star is not None and int(star) <= 2:
        level = "强负"
    elif star == 3:
        level = "负面"
    elif star == 5 and (r.get("has_image") or len(text) >= 15):
        level = "强正"
    else:
        level = "正面"

    urgent = bool(re.search(URGENT_WORDS, text))
    if level in ("强负", "负面") and r.get("has_image"):
        pri = "P0"
    elif level == "强负":
        pri = "P1"
    elif level == "负面":
        pri = "P2"
    else:
        pri = "好评"
    if level in ("强负", "负面") and r.get("is_followup") and pri != "P0":
        pri = {"P1": "P0", "P2": "P1"}[pri]

    if pri in ("P0", "P1", "P2") and urgent:
        action = "转人工+店长介入（涉诉风险词），留证原话与截图"
    elif pri == "P0":
        action = "安抚回复（appease-script-generate）+ 图证留存；质量类同步排查批次"
    elif pri == "P1":
        action = "安抚回复 + 根因归因（negative-review-rootcause）"
    elif pri == "P2":
        action = "标准回复 + 期望管理口径"
    else:
        action = "好评回复 30-80 字（不诱导晒图返利）"
    return level, pri, action


def build(payload, outdir):
    shop = payload.get("shop", "未提供")
    platform = payload.get("platform", "未提供")
    date = payload.get("monitor_date", "")
    reviews = payload.get("reviews", [])

    todo, good_replies, human_case = [], [], []
    cnt = Counter()
    for i, r in enumerate(reviews, 1):
        level, pri, action = classify(r)
        cnt[pri] += 1
        row = {
            "序号": i, "星级": r.get("star", "—"),
            "分级": level, "优先级": pri,
            "SLA": f"{SLA_HOURS[pri]} 小时内",
            "处置动作": action,
            "追评": "是" if r.get("is_followup") else "否",
            "摘要": str(r.get("text", ""))[:28] + ("…" if len(str(r.get("text", ""))) > 28 else ""),
        }
        if pri == "好评":
            good_replies.append(row)
        else:
            todo.append(row)
        if "转人工" in action:
            human_case.append(f"#{i} {row['摘要']}")

    todo.sort(key=lambda x: {"P0": 0, "P1": 1, "P2": 2}[x["优先级"]])
    backlog = [t for t in todo if t["优先级"] == "P0"]
    alert = (f"⚠️ 今日 P0 未清零：{len(backlog)} 条带图差评须 {SLA_HOURS['P0']} 小时内处置"
             if backlog else "✅ 今日无 P0 积压")

    summary = {
        "店铺": shop, "平台": platform, "监控日期": date,
        "评价总数": len(reviews),
        "优先级分布": dict(cnt),
        "P0": cnt.get("P0", 0), "P1": cnt.get("P1", 0), "P2": cnt.get("P2", 0),
        "转人工": len(human_case),
        "今日告警": alert,
        "SLA 规则": f"P0={SLA_HOURS['P0']}h P1={SLA_HOURS['P1']}h P2={SLA_HOURS['P2']}h 好评={SLA_HOURS['好评']}h",
        "编排说明": ("S1 校验 → S2 分级(同源 review-sentiment-analyze 规则) → S3 SLA 排期 → "
                     "S4 处置分发；话术细写由模型按 prompt.txt 完成，数值以脚本输出为准"),
    }

    at.ensure_outdir(outdir)
    xlsx = at.write_excel(
        os.path.join(outdir, "评价监控处置清单.xlsx"),
        {
            "今日必办": todo or [{"序号": "", "摘要": "（今日无差评待办）"}],
            "好评回复": good_replies,
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"今日必办": {"优先级": "contains:P0"}},
        widths={"今日必办": {"处置动作": 46, "摘要": 32}},
    )
    sections = [
        {"heading": "一、今日告警", "paras": [alert, f"SLA 规则：P0 {SLA_HOURS['P0']} 小时 / P1 {SLA_HOURS['P1']} 小时 / P2 {SLA_HOURS['P2']} 小时 / 好评 {SLA_HOURS['好评']} 小时内回复。"]},
        {"heading": "二、转人工名单（涉诉风险）",
         "bullets": human_case or ["（无）"], "paras": ["涉诉案件停用模板话术，店长 2 小时内介入，留证原话与截图。"]},
        {"heading": "三、今日必办（按 P0→P2 排序）",
         "table": {"cols": ["序号", "优先级", "SLA", "分级", "摘要", "处置动作"],
                   "rows": [[t["序号"], t["优先级"], t["SLA"], t["分级"], t["摘要"], t["处置动作"]] for t in todo]}},
        {"heading": "四、好评回复队列", "paras": [f"共 {len(good_replies)} 条，24 小时内逐条 30-80 字回复，不诱导晒图返利。"]},
        {"heading": "五、人工确认事项",
         "bullets": ["转人工案件的补偿口径由店长确认；对外回复发布前人工审核",
                     "本清单为 AI 生成内容，优先级与数值以脚本输出为准"]},
    ]
    docx = at.write_docx(
        os.path.join(outdir, "评价监控处置清单.docx"),
        f"{shop} 评价监控处置清单（{date}）", sections,
        subtitle=f"{platform} · AI 生成内容",
    )
    js = at.write_json({"summary": summary, "todo": todo, "good_replies": good_replies,
                        "generated_at": at.stamp(),
                        "note": "机器分级与排期；话术与语境由模型按 prompt.txt 复核"},
                       os.path.join(outdir, "monitor_flow.json"))
    return {"files": [xlsx, docx, js], "summary": summary, "alert": alert}


def main():
    ap = argparse.ArgumentParser(description="评价流实时监控与分级")
    ap.add_argument("--input", help="输入 JSON（shop/platform/monitor_date/reviews）")
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
    print(r["alert"])
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
