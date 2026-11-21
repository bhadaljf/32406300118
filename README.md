# 面向对象分析与设计 实验报告

广州大学《面向对象分析与设计》课程实验，对应《实验指导书》V202610。

| 实验 | 位置 | 内容 |
|---|---|---|
| 实验一 需求分析 | 仓库根目录 | 手工建模 + AI 四活动，PDF 32 页 |
| 实验二 系统设计 | `实验二-系统设计/` | 4+1 视图设计 + 两轮提示词 + AI 备选方案论证，PDF 26 页 |
| 实验三 综合实验 | `实验三-综合实验/` | 全系统分析、设计、原型与验证，PDF 21 页 |

## 实验一：需求分析

校园实验室预约管理系统需求分析实验报告，覆盖 V202610 实验一的四个活动。

### 目录结构

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
| `evidence/` | AI 对话截屏 |

## 生成方式

```
pwsh build/render_diagrams.ps1        # 需要 java 与 plantuml.jar
python build/build_report.py
python build/fix_heading_table_format.py
pwsh build/finalize.ps1               # 需要本机 Word
```

`render_diagrams.ps1` 默认指向 VS Code PlantUML 插件自带的 plantuml.jar，可用 `-PlantUmlJar` 指定其他路径。

注意：`out/` 下的报告曾按手工意见修订，**不要再用构建脚本覆盖**；只运行 `finalize.ps1` 重新导出 PDF。

## 交付清单

| 交付物 | 位置 |
|---|---|
| 实验一报告正文 | `out/面向对象分析与设计实验1-软件243_封媛.docx` |
| 实验一报告 PDF | `out/面向对象分析与设计实验1-软件243_封媛.pdf`，32 页 |
| 手工模型源文件 | `diagrams/01`–`04` 的 `.puml` |
| AI 生成件源文件 | `diagrams/05`、`06` 的 `.puml` |
| Mermaid 对照源文件 | `diagrams/08`、`09` 的 `.mmd` |
| 实验二报告 | `实验二-系统设计/out/面向对象分析与设计实验2-软件243_封媛.pdf`，26 页 |
| 实验二模型源文件 | `实验二-系统设计/diagrams/*.puml`（7 张） |
| 实验三报告 | `实验三-综合实验/out/面向对象分析与设计实验3-软件243_封媛.pdf`，21 页 |
| 实验三模型源文件 | `实验三-综合实验/diagrams/*.puml`（9 张） |

## 说明

- 实验一活动 1 的模型内容由本人确定，未使用 AI 生成；活动 2 至活动 4 使用 DeepSeek 在线对话生成对照模型，产出仅作为被评估对象。
- 实验二覆盖四个活动：手工设计、两轮提示词、AI 备选方案（两套体系结构 + 两版类图 + 并发顺序图）、三维度论证与缺陷清单。
- 实验三选题为自拟（延续本系统），需按课程要求报教师备案；原型为设计原型，实现范围见报告 4.4 节，建表语句与接口样例见附录 B。
- 实验三的 AI 四活动中，活动 1 已完成，活动 2 提示词见第 11 章，活动 3、4 待补充。
- 全部模型以 PlantUML 文本维护，报告内嵌渲染图，不作仅截图交付。
