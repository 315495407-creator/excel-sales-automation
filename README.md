# 练习 01：合并多个 Excel 销售文件

## 目标

把 `input` 文件夹中的多个 `.xlsx` 销售文件合并成一个报表，输出到 `output/merged_sales_report.xlsx`，同时生成处理日志。

## 运行前准备

在 `input` 文件夹放入一个或多个 Excel 文件。每个文件至少包含这些列：

- 日期
- 销售员
- 产品
- 数量
- 金额

原始文件不要放到 `output` 文件夹，也不要直接覆盖原始数据。

## 运行命令

```powershell
python merge_sales_files.py
```

如果使用本机的 Python 运行失败，请改用工作台提供的 Python 路径运行。

## 你需要观察的结果

1. `output/merged_sales_report.xlsx` 是否生成；
2. 合并后的行数是否等于各输入文件有效行数之和；
3. `output/process.log` 是否记录了每个文件的处理结果；
4. 故意放入一个缺少列的文件，观察脚本如何处理错误。

## 学习任务

- 先通读脚本，不急着修改；
- 找出读取 Excel、合并数据、保存 Excel 的三行代码；
- 修改必需列，加入你自己的字段；
- 为报表增加“总金额”和“平均金额”的统计信息。
