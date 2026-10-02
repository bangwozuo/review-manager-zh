# 测试报告

## 测试环境

- Python：`C:\Users\nsxzy\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
- 依赖：openpyxl / python-docx

## 实跑记录（2026-09-30）

| 命令 | 退出码 | 结果 |
|---|---|---|
| `python scripts/run_flow.py --demo` | 0 | 6 条差评 → L1×2 / L2×1 / L4×1 / L3×2；人工确认 1 件 |
| `python scripts/run_flow.py --input examples/input.json --outdir out` | 0 | 与 demo 同源数据，结果一致 |

## 产物清单（实得文件）

| 文件 | 大小 | 说明 |
|---|---|---|
| out/安抚补偿方案.xlsx | 8.7 KB | 分级方案（L5 标红）/ 授权核对 / 汇总 |
| out/安抚补偿方案.docx | 37.4 KB | 交接 Word（人工确认清单 + 执行提醒） |
| out/appease_flow.json | 4.1 KB | 机器可读结果 |

## 分步验证

| 步骤 | 验证点 | 结果 |
|---|---|---|
| S1 筛查 | 根因判定（物流/安全/工艺/期望差） | 通过 ✅ |
| S2 分级 | L4「茶叶异物」命中安全词；成本上限 30%/60% 客单价 | 通过 ✅ |
| S3 话术 | 30-80 字四件套；L4 走店长专线口径 | 通过 ✅ |
| S4 核对 | 「无全额退款授权」被正确识别为超授权 | 通过 ✅ |

## 缺陷修复记录

| 缺陷 | 修复 | 复跑 |
|---|---|---|
| policy「无全额退款授权」因子串匹配被误判为已授权 | 改为排除「无全额退款」前缀的判定 | 通过 ✅（L4 正确进人工确认清单） |
| Excel 授权核对 sheet 出现 dict 值导致 openpyxl 报错 | 统一 str() 转换 | 通过 ✅ |

---

*本报告为真实实跑证据，产物文件随仓库交付。*
