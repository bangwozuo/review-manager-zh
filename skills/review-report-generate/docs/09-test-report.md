# 测试报告

## 测试环境

- Python：`C:\Users\nsxzy\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
- 依赖：python-docx / openpyxl / matplotlib

## 实跑记录（2026-09-30）

| 命令 | 退出码 | 结果 |
|---|---|---|
| `python scripts/report_gen.py --demo` | 0 | 本周评分 4.62 低于 4.7 达标线；首要根因物流占 56% |
| `python scripts/report_gen.py --input examples/input.json --outdir out` | 0 | 与 demo 同源数据，结果一致 |

## 产物清单（实得文件）

| 文件 | 大小 | 说明 |
|---|---|---|
| out/评价周报.docx | 78.1 KB | 六段式 Word 周报（评分走势图内嵌） |
| out/评价周报数据.xlsx | 8.1 KB | 指标对照（未达标标红）/ 根因分布 / 改进项 |
| out/评分走势.png | 43.7 KB | 7 天评分折线图 |
| out/report.json | 2.4 KB | 机器可读结果 |

## 关键判定核对

| 项 | 期望 | 实测 | 结论 |
|---|---|---|---|
| 评分 4.62 vs 4.7 | 未达标 | 未达标 ✅ | 通过 |
| 差评响应 3.5h vs 2h | 未达标 | 未达标 ✅ | 通过 |
| 一般响应 21h vs 24h | 达标 | 达标 ✅ | 通过 |
| 根因 Top3 改进项 | 带数字目标+责任人+时点 | 全带 ✅ | 通过 |
| Word 表格与内嵌图 | 六段齐全 | 齐全 ✅ | 通过 |

## 边界用例

- 缺 `last_metrics` → 上周列「—」，不崩不估算
- 缺 `score_trend` → 跳过走势图，其余产物正常
- 缺 `root_causes` → 改进项区显示「（本周无差评）」

---

*本报告为真实实跑证据，产物文件随仓库交付。*
