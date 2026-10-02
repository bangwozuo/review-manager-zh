# 测试报告

## 测试环境

- Python：`C:\Users\nsxzy\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
- 依赖：openpyxl / matplotlib（assettools 按需检查，缺失时打印修复命令退出码 2）

## 实跑记录（2026-09-30）

| 命令 | 退出码 | 结果 |
|---|---|---|
| `python scripts/sentiment_analyze.py --demo` | 0 | 20 条评价：好评率 50.0% / 差评率 35.0%，P0 带图差评 4 条 |
| `python scripts/sentiment_analyze.py --input examples/input.json --outdir out` | 0 | 同上（与 demo 同源数据交叉验证一致） |

## 产物清单（实得文件）

| 文件 | 大小 | 说明 |
|---|---|---|
| out/评价情感分析.xlsx | 9.8 KB | 分级明细 20 行（P0 标红）/ 标签词频 / 健康度对照 / 汇总 |
| out/情感分布.png | 24.5 KB | 五级情感占比饼图 |
| out/负面标签Top.png | 26.3 KB | 负面标签加权频次柱状图 |
| out/sentiment.json | 7.7 KB | 机器可读结果 |

## 关键命中明细（抽样核对）

| 输入 | 期望 | 实测 | 结论 |
|---|---|---|---|
| 星 1 + 带图 + 12315 | P0 + 转人工 | P0 / 转人工 ✅ | 通过 |
| 星 3 + 追评 | 升级 P1 | P1 ✅ | 通过 |
| 星 5 + 文字 2 字 | 正面（非强正） | 正面 ✅ | 通过 |
| 好评率基准 | < 95% 判未达标 | 50.0% 判须提升 ✅ | 通过 |

## 边界用例

- 全部评价无文字且无星级 → 全部判「中」，不报错
- 空列表 → 差评率/好评率按 0 样本兜底（n=1 防除零），汇总正常产出

---

*本报告为真实实跑证据，产物文件随仓库交付。*
