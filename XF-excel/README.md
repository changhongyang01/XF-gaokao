# XF-excel

一键生成完整高考志愿 Excel 工作簿。

## 输出内容

- **Sheet 1**: 冲稳保总览（红黄绿三色标注）
- **Sheet 2**: 近三年趋势 + 安全评估
- **Sheet 3**: 院校详情 + 优劣分析
- **Sheet 4**: 四年费用概算 + ROI
- **图表**: 分数走势柱状图 + 你的位置线

## 示例

> 上海考生政史地538分，把这些志愿生成Excel对照表

## 使用

提供 JSON 数据 → 运行 `scripts/generate.py` → 输出 `.xlsx` 文件

## 联动

从 XF-gaokao / XF-rational / gaokao-admission-table 获取数据后组装 JSON 调用。