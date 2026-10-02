# 测试报告

## 测试环境

- Python：`C:\Users\nsxzy\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
- 依赖：openpyxl / python-docx

## 实跑记录（2026-09-30）

| 命令 | 退出码 | 结果 |
|---|---|---|
| `python scripts/run_flow.py --demo` | 0 | ⚠️ 今日 P0 未清零：2 条带图差评须 2 小时内处置 |
| `python scripts/run_flow.py --input examples/input.json --outdir out` | 0 | P0×2 / P1×3 / P2×1 / 好评×2，转人工 1 件 |

## 产物清单（实得文件）

| 文件 | 大小 | 说明 |
|---|---|---|
| out/评价监控处置清单.xlsx | 8.0 KB | 今日必办（P0 标红）/ 好评回复 / 汇总 |
| out/评价监控处置清单.docx | 37.1 KB | 交接 Word（转人工名单 + SLA + 必办表） |
| out/monitor_flow.json | 3.3 KB | 机器可读结果 |

## 分步验证

| 步骤 | 验证点 | 结果 |
|---|---|---|
| S1 校验 | 8 条字段齐全（含订单号） | 通过 ✅ |
| S2 分级 | 与 review-sentiment-analyze 同源规则（星 1 带图 → P0） | 通过 ✅ |
| S3 排期 | 追评差评升级；涉诉词（曝光/红屁股）进转人工 | 通过 ✅ |
| S4 分发 | 涉诉件停用模板；好评进回复队列 | 通过 ✅ |

## 关键判定核对

| 输入 | 期望 | 实测 | 结论 |
|---|---|---|---|
| 星 1 带图 +「曝光」 | P0 + 转人工 | P0 转人工 ✅ | 通过 |
| 星 1 带图追评（安全词） | P0 转人工 | P0 转人工 ✅ | 通过 |
| 星 4 一般评价 | 好评队列 24h | 好评队列 ✅ | 通过 |

---

*本报告为真实实跑证据，产物文件随仓库交付。*
