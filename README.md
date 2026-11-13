# 面向对象分析与设计 实验一：需求分析

校园实验室预约管理系统需求分析实验报告。

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

## 版本

对应《面向对象分析与设计实验指导书》V202610。全部模型以 PlantUML 文本维护，不作仅截图交付。
