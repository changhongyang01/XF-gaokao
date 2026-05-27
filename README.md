# Codex Gaokao Skills

三个 Codex Skill，专为高考志愿填报场景打造。

## Skill 清单

| Skill | 说明 |
|-------|------|
| XF-excel | 一键生成完整 Excel 工作簿——四张 Sheet + 柱状图，红黄绿冲稳保色标 |
| XF-rational | 六张底牌理性决策——录取、就业、地域、成本、匹配、长期，全维度分析 |

| 内置数据 | XF-gaokao/references/admission-data.md — 近5年全国录取数据参考 + 31省分数体系总览（总分/科目/赋分/分段分布）（批次线、院校位次、艺术类公式） |

| `XF-gaokao` | XF风格志愿填报指导，就业导向，直白话术 |
| `gaokao-admission-table` | 近三年录取数据对照表（Markdown + Excel），冲稳保评估，含学费对比 |

## 安装

将 `XF-gaokao/` 和 `gaokao-admission-table/` 两个文件夹复制到 Codex skills 目录：

```
C:\Users\你的用户名\.codex\skills\
```

重启 Codex 或新建对话即可生效。

## 使用建议

两个 Skill 搭配使用效果最好：

1. 先用 `gaokao-admission-table` 生成目标院校近三年录取数据对比表
2. 再用 `XF-gaokao` 根据数据给出冲稳保策略和最终建议

## 示例

> 上海考生政史地538分，帮我做一份完整志愿填报方案

> 上海考生政史地538分，帮我做一份完整志愿填报方案含Excel导出

## 许可

MIT License

## 声明

- 分数线数据具有时效性，请以当年各省教育考试院公布为准
- 建议仅供参考，最终决定请结合个人实际情况
- 非XF本人出品，仅为风格化 AI 指导