# -*- coding: utf-8 -*-
"""
平台规则问答 —— 本地规则库检索器（关键词打分召回 + 条款引用）。

职责边界：本脚本只做**内置规则库的关键词打分检索、条款召回与产物生成**（机器的强项）。
条款在具体案件中的适用判断、申诉话术撰写由模型按 prompt.txt 完成（这是模型的强项）。

内置规则库范围（评价运营高频条款，来源为各平台公开规则要点，以官方最新公示为准）：
  恶意差评申诉（证据要求/受理范围）、评价删除与修改规则、诱导好评认定与处罚、
  刷单炒信处罚（《反不正当竞争法》第 8 条 + 平台细则）、回复率考核标准、
  催评合规边界、差评骚扰认定、退款售后时效。

检索规则（量化标准）：
  query 分词后与规则的关键词表逐词匹配：命中关键词 ×2、命中标题 ×3、命中正文 ×1
  得分 ≥ 3 分召回；返回 Top 3；最高分 < 3 → 判定「知识库未覆盖」并给出缺口清单
  恶意差评类 query 附带证据包自查清单（订单记录/聊天记录/物流凭证）

用法：
  python rules_qa.py --input input.json --outdir out
  python rules_qa.py --demo

产物：
  out/规则检索结果.xlsx   检索问答 / 命中条款 / 汇总
  out/rules_qa.json       机器可读结果（供工作流读取）
"""
from __future__ import annotations

import argparse
import os
import re
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

RECALL_THRESHOLD = 3   # 总分 ≥ 3 召回，最高分 < 3 判定知识库未覆盖
TOP_K = 3              # 每个问题最多召回 3 条条款

# 内置规则库。来源：淘宝/天猫/京东/拼多多公开评价与售后规则要点（以官方最新公示为准）。
RULES = [
    {
        "id": "R01", "title": "恶意差评申诉：受理范围与证据要求",
        "keywords": ["恶意差评", "差评", "申诉", "投诉", "证据", "举报", "同行", "敲诈"],
        "content": ("各平台均设恶意行为投诉通道（天猫-恶意行为投诉中心、京东-商家客服 jealous、"
                    "拼多多-商家后台举报）。可申诉类型：同行攻击（订单记录无关联）、敲诈勒索"
                    "（聊天记录里要钱/要赠品换改评）、涉政辱骂、无实际订单的评价。"
                    "证据包三件套：①订单记录（证明真实交易或证明无交易）②聊天记录（完整会话，"
                    "平台要求原始截图含时间戳）③物流凭证（已拒收/未签收面单）。"
                    "申诉成功率行业经验约 40-60%；证据不全（缺聊天记录或订单号对不上）不要提交，"
                    "会被驳回且占用当日申诉次数。"),
        "source": "各平台商家后台恶意行为投诉规则要点",
    },
    {
        "id": "R02", "title": "评价删除与修改：平台规则",
        "keywords": ["删除", "删差评", "修改", "改评", "评价", "撤销", "屏蔽"],
        "content": ("淘宝/天猫：评价作出后 30 天内买家可自行修改或删除一次，商家无权限删除买家评价；"
                    "拼多多：买家评价无法自行修改，只能追评一次补充；京东：晒单评价删除后不恢复。"
                    "商家侧唯一合规路径：①联系买家自然沟通解决问题（不得利诱）②恶意差评走申诉。"
                    "任何「付费删差评」「差评屏蔽服务」均为违规宣称，涉嫌违法。"),
        "source": "淘宝评价规则 / 拼多多评价管理规范要点",
    },
    {
        "id": "R03", "title": "诱导好评与返现：认定与处罚",
        "keywords": ["好评返现", "返现", "返红包", "诱导", "好评卡", "换好评", "利诱", "催评"],
        "content": ("好评返现、好评返红包、好评送赠品、评论有礼（限定好评）均属诱导好评，"
                    "违反《反不正当竞争法》第 8 条虚假宣传，平台处罚：删除违规评价、商品降权、"
                    "扣分（天猫一般违规每次扣 6 分，累计 48 分封店）、保证金扣除。"
                    "合规催评方式：订单完成后包裹卡/短信中性提醒「如有问题联系客服」，"
                    "不得承诺任何利益换好评，不得出现「五星好评+返现」字样。"),
        "source": "《反不正当竞争法》第 8 条 / 天猫淘宝评价规则",
    },
    {
        "id": "R04", "title": "刷单炒信：法律与平台双重处罚",
        "keywords": ["刷单", "刷好评", "刷评价", "炒信", "虚假交易", "水军", "补单"],
        "content": ("刷单炒信违反《反不正当竞争法》第 8 条（虚假宣传），市场监管部门可处 20 万-100 万元罚款；"
                    "《电子商务法》第 17 条禁止虚构交易编造评价。平台处罚：虚假交易商品降权/下架、"
                    "店铺扣分、清退。买家账号也会被封禁。刷出来的好评集中在短期内相似文案率上升，"
                    "平台风控（如天猫蜜獾系统）识别后整批删除并追溯处罚。"),
        "source": "《反不正当竞争法》第 8 条 / 《电子商务法》第 17 条",
    },
    {
        "id": "R05", "title": "回复率与回复时长考核",
        "keywords": ["回复率", "回复", "回复时长", "响应", "考核", "答复"],
        "content": ("各平台对评价回复无强制考核，但回复率 100% 是运营基准：每条评价（含默认好评）"
                    "30-80 字个性化回复，避免「感谢惠顾」模板复制。回复时长目标：差评 2 小时内、"
                    "一般评价 24 小时内。回复本身公开展示，潜在买家会参考商家对差评的态度——"
                    "专业回复可挽回 5%-10% 的转化犹豫。禁止在回复中留联系方式或引导站外交易。"),
        "source": "评价运营行业基准 / 各平台评价回复规范",
    },
    {
        "id": "R06", "title": "差评骚扰与报复买家认定",
        "keywords": ["骚扰", "报复", "电话", "短信", "轰炸", "人身威胁"],
        "content": ("商家或雇佣第三方电话短信轰炸买家、恐吓报复差评买家，构成平台「骚扰他人」违规"
                    "（天猫严重违规每次扣 48 分直接封店），且可能违反《治安管理处罚法》第 42 条。"
                    "合规联系买家：订单详情页留言、平台 IM 一次、每天不超过 2 次沟通尝试，"
                    "买家明确拒绝沟通后停止。"),
        "source": "天猫/淘宝平台规则「骚扰他人」条款要点",
    },
    {
        "id": "R07", "title": "退款售后时效规则",
        "keywords": ["退款", "退货", "售后", "七天无理由", "时效", "运费险"],
        "content": ("七天无理由退货：自签收次日起 7 天，商品完好可退，生鲜/定制/贴身衣物除外；"
                    "质量问题退货运费由商家承担，无理由退货由买家承担（运费险除外）。"
                    "退款响应：商家须在 48 小时内响应买家退款申请，超时系统默认同意。"
                    "质量问题争议以商品照片+检测报告定责。"),
        "source": "《消费者权益保护法》第 25 条 / 各平台售后规则",
    },
    {
        "id": "R08", "title": "评价展示与权重规则",
        "keywords": ["追评", "展示", "权重", "置顶", "排序", "带图"],
        "content": ("评价展示权重：带图评价 > 文字评价 > 默认好评；追评比首评权重更高"
                    "（收货使用后的补充评价更可信），追评差评对转化的杀伤力大于首评差评。"
                    "负面评价置顶规则：多数平台按「有用数 + 时间衰减」排序，差评一旦获得点赞"
                    "会长期占据首屏，因此带图差评必须 2 小时内响应处置。"),
        "source": "各平台评价排序规则要点 / 电商运营行业共识",
    },
]

EVIDENCE_CHECKLIST = [
    "订单记录（真实交易证明或无交易证明）",
    "完整聊天记录截图（含时间戳，未经裁剪）",
    "物流凭证（签收/拒收/未发货面单）",
    "评价原文截图（含发布时间与账号）",
    "买家历史评价行为（是否惯发性恶意差评）",
]

DEMO = {
    "queries": [
        "买家给了恶意差评还能申诉吗？要准备什么证据？",
        "好评返现违规吗？会被怎么处罚？",
        "差评可以花钱删除吗",
    ],
}


def tokenize(query: str) -> list:
    """粗分词：先按规则库关键词做包含匹配的「词表投影」，再切 2-gram 兜底。"""
    vocab = set()
    for r in RULES:
        vocab.update(r["keywords"])
    hits = [w for w in vocab if w in query]
    return hits or [query[i:i + 2] for i in range(len(query) - 1)]


def retrieve(query: str):
    toks = tokenize(query)
    scored = []
    for r in RULES:
        s = 0
        for t in toks:
            if t in r["keywords"]:
                s += 2
            if t in r["title"]:
                s += 3
            if t in r["content"]:
                s += 1
        if s >= RECALL_THRESHOLD:
            scored.append({"规则ID": r["id"], "条款标题": r["title"], "匹配得分": s,
                           "内容要点": r["content"], "出处": r["source"]})
    scored.sort(key=lambda x: -x["匹配得分"])
    return scored[:TOP_K]


def build(payload, outdir):
    queries = payload.get("queries", [])
    if isinstance(queries, str):
        queries = [queries]

    qa_rows, all_hits, gaps = [], [], []
    for qi, q in enumerate(queries, 1):
        hits = retrieve(str(q))
        if hits:
            best = hits[0]
            qa_rows.append({
                "#": qi, "问题": q, "召回条款": f"{best['规则ID']} {best['条款标题']}",
                "匹配得分": best["匹配得分"],
                "回答要点": best["内容要点"][:120] + "…",
                "出处": best["出处"],
            })
            for h in hits:
                all_hits.append({"问题#": qi, "问题": q, **h})
        else:
            qa_rows.append({"#": qi, "问题": q, "召回条款": "（知识库未覆盖）",
                            "匹配得分": 0, "回答要点": "列入知识缺口，需人工补充官方规则原文",
                            "出处": "—"})
            gaps.append(q)
        if re.search(r"恶意差评|申诉|差评.*证据", str(q)):
            qa_rows[-1]["回答要点"] += " ‖ 证据包自查：" + "；".join(EVIDENCE_CHECKLIST[:3])

    summary = {
        "问题数": len(queries),
        "成功召回数": sum(1 for r in qa_rows if r["匹配得分"] >= RECALL_THRESHOLD),
        "知识缺口数": len(gaps),
        "召回阈值": RECALL_THRESHOLD,
        "规则库条款数": len(RULES),
        "说明": ("脚本只做关键词打分检索与条款召回，召回阈值 3 分、返回 Top 3；"
                 "条款在具体案件中的适用判断由模型按 prompt.txt 完成；"
                 "平台规则以官方最新公示为准"),
    }

    at.ensure_outdir(outdir)
    xlsx = at.write_excel(
        os.path.join(outdir, "规则检索结果.xlsx"),
        {
            "检索问答": qa_rows,
            "命中条款明细": all_hits or [{"问题#": "", "问题": "", "规则ID": "", "条款标题": "", "匹配得分": 0, "内容要点": "", "出处": ""}],
            "申诉证据包清单": [{"#": i, "证据项": e, "是否齐全": "待人工核对"} for i, e in enumerate(EVIDENCE_CHECKLIST, 1)],
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"检索问答": {"召回条款": "contains:未覆盖"}},
        widths={"检索问答": {"问题": 30, "回答要点": 50}, "命中条款明细": {"内容要点": 50}},
    )
    js = at.write_json({"summary": summary, "qa": qa_rows, "gaps": gaps,
                        "generated_at": at.stamp(),
                        "note": "检索结果；条款适用判断由模型按 prompt.txt 复核"},
                       os.path.join(outdir, "rules_qa.json"))
    return {"files": [xlsx, js], "summary": summary}


def main():
    ap = argparse.ArgumentParser(description="平台规则问答 —— 本地规则库检索")
    ap.add_argument("--input", help="输入 JSON（queries: [...]）")
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
    print(f"{s['问题数']} 个问题：召回 {s['成功召回数']} 个，知识缺口 {s['知识缺口数']} 个")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
