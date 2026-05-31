---
name: docx-code
description: >
  Generate project documentation (design specs, function/feature manuals) as a .docx
  from a codebase, with all formatting strictly inherited from an existing .docx template.
  Use when the user wants a Word document that describes a software project's modules and
  features based on its real source code, while reusing a reference document's fixed styles
  for headings, body text, figures and tables. 依据代码仓库生成项目文档（设计说明书/功能说明书），
  并严格沿用模板 .docx 中已有的固定样式，统一各级标题、正文、图、表格式。
---

# docx-code：依据代码生成格式统一的项目文档

## 适用场景

需要把一个软件项目（可含前端/后端/算法多个仓库）的**真实功能**整理成一份正式 Word 文档
（设计说明书、功能说明书等），且**格式必须与一份已有的标准 .docx 模板严格一致**时使用。

核心做法：复制模板 .docx，完整继承其样式定义，**只重写正文**，正文一律引用模板**已有的固定样式**——
因此各级标题、正文、图、表的字体/字号/缩进/行距/编号都与模板完全统一，不自造任何样式。

引擎 `scripts/docxgen.py` 仅用 Python 标准库，无需 python-docx / pandoc / LibreOffice，
可在任意装有 Python 3 的系统直接运行。

## 前置条件

1. 一份**格式标准的参考 .docx 模板**（提供全部样式与版式）。
2. 待编写文档的**源代码仓库**。
3. Python 3（标准库即可）。

## 工作流

### 1. 解析模板，摸清可用样式与层级范式
```
python scripts/inspect_template.py <模板.docx>
```
输出 docxgen 将使用的样式映射（各级标题 / 正文 / 题注 / 表格的 styleId、字体、字号）与模板现有大纲。
据此理解模板的**层级范式**（例如：功能模块[H1] → 各大模块[H2] → 设计模块[H3] → 模块功能组成[H4] → 子模块[H5]），
新文档应沿用同一层级模式。

### 2. 分析代码仓库，提炼功能模块树
基于真实代码（**不臆测**）梳理"一级功能模块 → 子功能"：
- **前端**：路由配置、菜单、`views` 目录结构、关键页面组件 → 功能模块与子功能划分
- **后端**：Controller/API 分组、Service、领域实体/枚举 → 能力与端点
- **算法**：模块目录、各算法实现 → 方法/算子清单

大型或多仓库项目**用子代理并行分析**（每个仓库一个 Explore/general 代理），统一要求输出：
`## 一级模块：X` 下逐条 `- 子功能 — 一句话作用 [证据: 路径]`，并交叉印证。
若目标项目与模板属同类系统，模板中**一致的功能模块结构可直接复用**，差异集中在核心算法/业务模块。

### 3. 设计文档大纲
把功能模块树映射到模板层级（H1 总章 → H2 大模块 → H3 设计模块 → H4 功能组成 → H5 子模块）。
功能模块作为一级标题，具体功能归入其下。CRUD/导入导出类展开到子模块层；并列的算法/方法各占一节。

### 4. 用 docxgen 生成
把大纲与正文写成一个**生成脚本**（便于迭代重跑），顺序调用 `b.h/b.p/b.table/b.image`，最后 `b.save()`：
```python
import sys; sys.path.insert(0, 'scripts')
from docxgen import DocxBuilder

with DocxBuilder('模板.docx') as b:       # 自动探测并继承模板样式
    b.h(1, '功能模块')
    b.p('本章描述各功能模块……')
    b.h(2, '样本数据接入模块')
    b.p('该模块负责……')
    b.h(3, '文件接入模块'); b.h(4, '模块功能组成')
    b.table(['序号', '格式', '说明'],
            [['1', 'Excel', '按表头解析'], ['2', 'CSV', '文本表格']],
            caption='支持的文件格式')
    b.h(5, 'Excel文件导入子模块'); b.p('支持导入 .xlsx……')
    # b.image('shots/login.png', caption='登录界面')   # 需要配图时
    print(b.save('输出.docx'))
```

### 5. 校验
`save()` 内置校验：XML 合法性 + 所有引用样式都存在于模板 + 返回结构统计
（`{'headings': {1:.., 2:..}, 'paragraphs': N, 'tables': M}`）。核对模块数与层级是否符合大纲。
可选版式预览（若已安装 LibreOffice，跨平台）：
```
soffice --headless --convert-to pdf 输出.docx
```

## 关键约定

- **只用模板已有的固定样式**（pStyle / tblStyle），不新增样式 → 全文格式天然统一。
- **标题多级编号**（1 / 1.1 / 1.1.1 …）由模板标题样式自带（outlineLvl + 多级列表），**不要手写编号**。
- **图/表题注**用 `SEQ` 域自动连续编号；表题在表上方、图题在图下方（docxgen 已内置）。
- **正文/表格**直接继承模板对应样式；表格附全框线、表头加粗+底纹、固定列宽。
- 内容描述基于真实代码，准确、信息密度高，避免空话套话。

## docxgen API 速查

| 方法 | 说明 |
|---|---|
| `DocxBuilder(template, styles=None, strip_media=True)` | 载入模板并探测样式；`strip_media=True` 清除模板自带图片得到纯文字+表格文档 |
| `.h(level, text)` | 第 level 级标题（用模板对应 styleId） |
| `.p(text, style=None)` | 正文段落（默认模板正文样式） |
| `.table(headers, rows, caption=None, widths=None)` | 表格；`caption` 生成表上方 SEQ 题注 |
| `.image(path, caption=None, width_emu=None)` | 嵌入 png/jpg，居中；`caption` 生成图下方 SEQ 题注 |
| `.raw(xml)` | 追加自定义合法 OOXML（高级） |
| `.save(out)` → dict | 校验并打包；返回结构统计。可重复调用 |
| `styles=` 覆盖 | `{'h':{1:'2',...}, 'body':'1', 'caption':'12', 'table':'23'}` |

详细工作流、子代理分工与完整示例见 [references/workflow.md](references/workflow.md)。
