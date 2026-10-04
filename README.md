# 评价运营官

> **把评价从售后成本变成可监控、可归因、可修复的运营资产**

[![Stage](https://img.shields.io/badge/stage-P0-orange)](https://github.com/bangwozuo)
[![Asset](https://img.shields.io/badge/asset-prompt%20%2B%20scripts-blueviolet)](#资产矩阵)
[![NoKey](https://img.shields.io/badge/API%20Key-not%20required-success)](#快速开始)
[![License](https://img.shields.io/badge/license-Apache--2.0-green)](LICENSE)

![演示](https://cdn.jsdelivr.net/gh/bangwozuo/review-manager-zh@main/docs/assets/hero.gif)

*▲ 实时演示（自动循环）· [▶ 观看完整版合集视频](https://cdn.jsdelivr.net/gh/bangwozuo/review-manager-zh@main/docs/demo.mp4)*

*上面 20 秒演示来自本仓 5 个代表资产的真实脚本执行截图：评价情感分析 → 差评根因分类 → 复盘报告生成 → 评价流监控 → 周报复盘（各资产完整截图见其 `docs/assets/run-terminal.png`）。*

---

## 它是谁

面向 **电商小卖家** 的数字员工资产包。

| 项目 | 内容 |
|------|------|
| 目标用户 | 退货率高、评分敏感类目（服饰/美妆/食品）一人店 |
| 身份边界 | 做：评价监控、根因归因、回复草案、证据整理；不做：联系买家删差评（违规）、刷单刷评 |
| KPI | 差评响应 ≤ 2 小时；中差评挽回率 ≥ 30%；DSR 环比不降 |
| 资产数 | 5 个技能 + 5 条工作流（9 个带确定性脚本，1 个纯提示词） |
| 旧名存档 | `评价运营专员·小评` |

---

## 资产矩阵

| 资产 | 一句话 | 类型 | README |
|------|--------|------|--------|
| 评价情感分析 | 五级情感分级 + P0/P1/P2 处置排期，早上 5 分钟知道先处理哪几条 | 技能 T1 | [README](skills/review-sentiment-analyze/README.md) |
| 差评根因分类 | 六类根因规则树归因，安全类 > 10% 当日下架自查 | 技能 T1 | [README](skills/negative-review-rootcause/README.md) |
| 平台规则问答 | 8 类条款库召回，申诉证据五项自查，不编条款编号 | 技能 T1 | [README](skills/platform-rules-qa/README.md) |
| 复盘报告生成 | 六段式周报 + 6 项指标基准对照，Word 一键落盘 | 技能 T1 | [README](skills/review-report-generate/README.md) |
| 安抚话术生成 | 30-80 字引用原文细节的差评回复，L1-L5 授权分级 | 技能 T2 | [README](skills/appease-script-generate/README.md) |
| 评价流实时监控与分级 | 当日评价流 →「今日必办」清单，P0 带图差评 2 小时 SLA | 工作流 T3 | [README](workflows/review-stream-monitor-flow/README.md) |
| 差评根因归因 | 差评池 → 治理排期表，噪声带 ±10pct，新发根因连业务事件 | 工作流 T3 | [README](workflows/negative-review-rootcause-flow/README.md) |
| 安抚回复与补偿方案生成 | 差评批次 → L1-L5 分级补偿方案册，超授权集中人工确认 | 工作流 T3 | [README](workflows/appease-compensation-flow/README.md) |
| 纠纷证据包整理 | 恶意差评证据五项打分，< 60% 禁止提交，申诉信四要素 | 工作流 T3 | [README](workflows/dispute-evidence-pack-flow/README.md) |
| 周报复盘 | 评价 + 指标 → 周会直接过的复盘 Word，改进 3 项带数字目标 | 工作流 T3 | [README](workflows/weekly-review-flow/README.md) |

---

## 资产形态

**提示词 + 确定性脚本** —— 理解本仓库的关键：

| 特性 | 说明 |
|------|------|
| ✅ 无需 API Key | 一个 Key 都不需要 |
| ✅ 双模式 | 纯提示词粘贴即用；脚本模式零 AI 依赖、确定性输出 Excel/Word/PNG |
| ✅ 产物真实 | 脚本实跑产物（xlsx/docx/png/json）随仓交付，README 截图来自真实执行 |
| ✅ 平台无关 | 提示词可粘贴到 Coze / WorkBuddy / Dify / Claude / ChatGPT |
| ✅ 用户自备算力 | 模型来自你自己的订阅 |

---

## 快速开始

**提示词方式**：

```text
1. 打开任一资产的 prompt.txt（如 skills/review-sentiment-analyze/prompt.txt）
2. 全文复制
3. 粘贴到你常用的 AI 工具
4. 按 SKILL.md 的输入规格提供数据
```

**脚本方式**（以情感分析为例）：

```bash
pip install -r requirements.txt
python skills/review-sentiment-analyze/scripts/sentiment_analyze.py --demo
```

完整指引见 [使用手册](docs/04-usage.md)。

---

## 仓库结构

```text
review-manager-zh/
├── README.md / employee.md / package.yaml     # 入口与 12 字段定义卡
├── docs/demo.mp4                              # 本仓演示录屏（真实执行截图串接）
├── docs/01~07                                 # 员工级文档（架构/流程/场景/手册/示例/录像/测试）
├── skills/                                    # 5 个原子技能
│   └── <skill>/
│       ├── README.md  SKILL.md  prompt.txt  schema.json  examples/
│       ├── scripts/                           # 确定性脚本（4/5 个技能有）
│       ├── out/                               # 脚本实跑产物（xlsx/png/json）
│       └── docs/                              # 该技能自己的 9 项文档 + run-terminal.png
├── workflows/                                 # 5 条工作流（复合技能，均带 run_flow.py）
├── knowledge/                                 # RAG wiki 知识库
├── connectors/                                # 连接器说明 + 合规红线
├── quality/                                   # 效果基线与追踪日志
└── tests/                                     # 资产校验测试（离线，无需密钥）
```

### 每个技能 / 工作流自带的 docs

| 文档 | 内容 |
|------|------|
| `README.md` | 资产速览（真实执行截图 + 规则表 + 真实 IO） |
| `docs/assets/run-terminal.png` | 真实执行终端截图 |
| `docs/01-usage-manual.md` | 安装使用手册 |
| `docs/02-architecture.md` | 业务架构图 |
| `docs/03-flow.md` | 流程图（Mermaid + 配图） |
| `docs/04-examples.md` | 使用示例 |
| `docs/06-scenarios.md` | 使用场景（适用 / 不适用） |
| `docs/08-value.md` | 解决问题与价值 |
| `docs/09-test-report.md` | 测试报告 |

---

## 交付物导航

| 文档 | 内容 |
|------|------|
| [业务架构](docs/01-architecture.md) | 四层架构 + 数据流 + 能力边界 |
| [工作流流程](docs/02-workflow.md) | 5 条工作流的 DAG 可视化 |
| [使用场景](docs/03-scenarios.md) | 3 个真实场景（含前后对比） |
| [使用手册](docs/04-usage.md) | 各平台导入指引 + 常见问题 |
| [示例库](docs/05-examples.md) | 5 组输入输出示例 |
| [录像脚本](docs/06-recording-script.md) | 7 镜头分镜 + 旁白稿 |
| [校验报告](docs/07-test-report.md) | 资产质量校验结果 |

---

## 知识库与连接器

| 目录 | 说明 |
|------|------|
| [`knowledge/`](knowledge/README.md) | RAG wiki 知识库：填入业务信息可显著提升输出质量 |
| [`connectors/`](connectors/README.md) | 连接器说明：数据从哪来、怎么合规地来 |

---

## 资产校验

```bash
pip install -r requirements.txt
pytest tests/ -v
```

校验技能完整性、提示词结构、契约一致性、工作流 DAG、技能级与工作流级 docs 完整性、知识库 wiki 与连接器结构。
**不需要任何 API Key。**

---

## 合规声明

- ✅ 所有输出为 **AI 辅助生成**，交付前须人工审核
- ✅ 提示词内置**违禁词禁止清单**，符合《广告法》《反不正当竞争法》要求
- ✅ 遵循《人工智能生成合成内容标识办法》
- ✅ 连接器只走**官方 API** 或**用户导出数据**，严禁爬取与诱导删评
- ✅ 所有对外发布动作**保留人工确认环节**

---

## 许可

[Apache-2.0](LICENSE) — 可自由使用、修改、商用

---

*由 bangwozuo 业务库自动生成 · 2026-09-29 · README 多媒体升级 2026-10-03*
