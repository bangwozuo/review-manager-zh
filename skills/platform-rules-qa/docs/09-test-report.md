# 测试报告

## 测试环境

- Python：`C:\Users\nsxzy\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
- 依赖：openpyxl（纯标准库检索逻辑 + Excel 产物）

## 实跑记录（2026-09-30）

| 命令 | 退出码 | 结果 |
|---|---|---|
| `python scripts/rules_qa.py --demo` | 0 | 3 个问题：召回 3 个，知识缺口 0 个 |
| `python scripts/rules_qa.py --input examples/input.json --outdir out` | 0 | 与 demo 同源数据，结果一致 |

## 产物清单（实得文件）

| 文件 | 大小 | 说明 |
|---|---|---|
| out/规则检索结果.xlsx | 10.9 KB | 检索问答（未覆盖标红）/ 命中条款明细 / 申诉证据包清单 |
| out/rules_qa.json | 2.6 KB | 机器可读结果 |

## 检索质量核对

| 问题 | 召回条款 | 得分 | 结论 |
|---|---|---|---|
| 恶意差评申诉证据 | R01 | 22 | 命中且自动附证据包清单 ✅ |
| 好评返现处罚 | R03 | 9 | 命中 ✅ |
| 花钱删差评 | R02 | 7 | 命中且给出负面合规口径 ✅ |

## 边界用例

- 无关问题（如「标题怎么优化」）→ 得分 < 3，标「知识库未覆盖」入缺口，不编造条款
- 单字符串 query（非数组）→ 自动包装为列表，正常运行
- 修复记录：一次 KeyError（召回字段「出处」键名不一致）已修复并复跑通过

---

*本报告为真实实跑证据，产物文件随仓库交付。*
