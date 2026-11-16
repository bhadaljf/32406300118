# 面向对象分析与设计 实验一：需求分析

校园实验室预约管理系统需求分析实验报告，对应《面向对象分析与设计实验指导书》V202610 实验一的四个活动。

## 目录结构

| 路径 | 说明 |
|---|---|
| `build/build_report.py` | 由模板生成报告正文 |
| `build/fix_heading_table_format.py` | 标题与表格格式后处理 |
| `build/finalize.ps1` | Word 更新目录域并导出 PDF |
| `build/render_diagrams.ps1` | 由 `.puml` 渲染模型图 |
| `build/render_pages.py` | 导出 PDF 分页图，用于版面校验 |
| `build/assets/实验一模板参考.docx` | 学校实验报告模板转换件 |
| `diagrams/` | 模型源文件（`.puml`）与渲染图（`.png`） |
| `out/` | 报告正文与最终 PDF |

## 生成方式

```
pwsh build/render_diagrams.ps1        # 需要 java 与 plantuml.jar
python build/build_report.py
python build/fix_heading_table_format.py
pwsh build/finalize.ps1               # 需要本机 Word
```

`render_diagrams.ps1` 默认指向 VS Code PlantUML 插件自带的 plantuml.jar，可用 `-PlantUmlJar` 指定其他路径。

## 交付清单

| 交付物 | 位置 |
|---|---|
| 实验报告正文 | `out/面向对象分析与设计实验1-软件243_封媛.docx` |
| 实验报告 PDF | `out/面向对象分析与设计实验1-软件243_封媛.pdf`，共 25 页 |
| 手工模型源文件 | `diagrams/01`–`04` 的 `.puml`（用例图、类图、状态图、顺序图） |
| AI 生成件源文件 | `diagrams/05`、`06` 的 `.puml`（活动 3 原始输出） |
| 渲染图 | `diagrams/*.png`，均已嵌入报告 |

## 报告结构

| 章节 | 内容 |
|---|---|
| 第 1–9 章 | 活动 1：手工建模，含用例字典、四类模型、追溯表与建模决策记录 |
| 第 10 章 | 两种建模方式的比较 |
| 第 11 章 | 活动 2 至活动 4：提示词设计、AI 原始输出、八项逐项对照、缺陷清单 D1–D8 |
| 附录 A | 手工模型源文件清单与提交哈希 |
| 附录 B | 第二轮 AI 原始输出 |

## 说明

- 活动 1 的模型内容由本人确定，未使用 AI 生成；活动 2 至活动 4 使用 DeepSeek 在线对话生成对照模型，产出仅作为被评估对象。
- 本报告以预约业务子系统（UC03–UC08）为分析对象。
- 全部模型以 PlantUML 文本维护，报告内嵌渲染图，不作仅截图交付。
