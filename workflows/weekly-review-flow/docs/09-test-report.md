# 测试报告

## 测试环境

- Python：`C:\Users\nsxzy\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
- 依赖：python-docx / openpyxl / matplotlib

## 实跑记录（2026-09-30）

| 命令 | 退出码 | 结果 |
|---|---|---|
| `python scripts/run_flow.py --demo` | 0 | 评分 4.66 < 4.7；差评率 4.1% > 3%；回复率 96% 漏约 12 条；连续 3 天下滑判趋势恶化 |
| `python scripts/run_flow.py --input examples/input.json --outdir out` | 0 | 与 demo 同源数据，结果一致 |

## 产物清单（实得文件）

| 文件 | 大小 | 说明 |
|---|---|---|
| out/评价周报复盘.docx | 79.2 KB | 六段式 Word 周报（走势图内嵌 + 人工确认事项） |
| out/周报复盘数据.xlsx | 8.7 KB | 指标对照（未达标标红）/ 根因分布 / 改进项 |
| out/情感与根因图.png | 44.8 KB | 7 天评分走势折线图 |
| out/weekly_flow.json | 2.8 KB | 机器可读结果 |

## 分步验证

| 步骤 | 验证点 | 结果 |
|---|---|---|
| S1 校验 | 12 条评价 + 指标齐全 | 通过 ✅ |
| S2 情感统计 | 与 review-sentiment-analyze 同源五级规则 | 通过 ✅ |
| S3 根因归因 | 与 negative-review-rootcause 同源；追评 ×2 | 通过 ✅ |
| S4 指标对照 | 六项基准（4.7/95%/3%/100%/2h/24h） | 通过 ✅ |
| S5 改进排期 | Top3 带数字目标 + 责任人 + 时点 | 通过 ✅ |

## 关键判定核对

| 输入 | 期望 | 实测 | 结论 |
|---|---|---|---|
| 连续 3 天下滑走势 | 趋势恶化告警 | 告警 ✅ | 通过 |
| 回复率 96% × 312 条 | 漏约 12 条 | 约 12 条 ✅ | 通过 |
| 差评率 4.1% | > 3% 专项治理 | 告警 ✅ | 通过 |
| Word 六段 + 内嵌图 | 齐全 | 齐全 ✅ | 通过 |

---

*本报告为真实实跑证据，产物文件随仓库交付。*
