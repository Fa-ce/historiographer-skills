---
name: docx-add-screenshots
description: 将浏览器截图批量插入 Word 文档的工具。适用于：给 docx 文档的指定章节/标题下方插入浏览器界面截图，每张截图带居中的"图 N"序号标注。截图通过 openCLI 获取，支持半自动模式（用户手动导航 → 自动截图插入）。当用户需要给 Word 文档添加界面截图、功能截图、操作截图时触发。跨平台：支持 Windows 原生和 WSL 环境。
---

# docx-add-screenshots

## 快速开始

```bash
python3 scripts/docx_screenshot_inserter.py --docx <文档路径> --opencli <opencli路径> [--session <session名>] [--width <宽>] [--height <高>]
```

## 工作流程

1. **备份**：自动创建 `.backup.docx`
2. **解包**：解析 docx 内部 XML 结构
3. **遍历目标**：找到所有匹配的标题段落
4. **逐个处理**（半自动）：
   - 显示当前标题 → 等待用户在浏览器中导航到目标界面
   - 用户按回车 → 自动 `bind` 绑定当前活跃 Chrome 标签页
   - 验证绑定的页面地址是否正确
   - 截图 → 插入文档（居中图片 + 下方"图 N"标注）
5. **打包**：重新生成 docx

## 截图流程（openCLI）

每次截图前自动执行：

1. `bind` — 绑定当前活跃 Chrome 标签页，获取 URL 并显示给用户确认
2. `screenshot` — 截取当前页面
3. 验证截图文件存在

**防错机制**：`bind` 后显示绑定的 URL，用户可在交互界面确认是否正确。如果绑定到错误窗口，用户可以输入 `r` 重新 bind。

## 图片格式规范

每张截图在文档中以两个段落呈现：

```
        [居中的嵌入式图片]          ← 图片段落，inline 嵌入，6.5" 宽度自适应
          图 N                      ← 图片标注，加粗居中
```

- **图片**：inline 嵌入（非浮动），宽度 6.5 英寸（标准页面可用宽度），高度按原始比例自动计算
- **标注**：格式为 `图 N`（N 从 1 递增），宋体加粗，居中对齐
- **单位**：使用 EMU（English Metric Units），1 英寸 = 914400 EMU

## 跨平台配置

脚本自动检测环境。用户需提供 openCLI 路径参数：

| 环境 | opencli 参数示例 | 说明 |
|---|---|---|
| WSL | `opencli.cmd` | 通过 powershell.exe 间接调用 |
| Windows 原生 | `opencli` 或 `opencli.cmd` | 直接调用 |
| 自定义命令 | `--opencli-cmd "full command"` | 完全自定义命令模板 |

### WSL 环境额外参数

```bash
--powershell /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe
```

## 脚本参数

```
--docx PATH          目标 docx 文件路径（必需）
--opencli PATH       opencli 可执行文件路径（必需）
--session NAME       openCLI session 名（默认: work）
--powershell PATH    powershell.exe 路径（WSL 环境需要）
--width NUM          截图宽度（默认: 1920）
--height NUM         截图高度（默认: 1080）
--output-dir PATH    截图保存目录（默认: 文档同级 screenshots/）
--heading-style VAL  目标标题的 style ID（默认自动检测所有 heading）
--heading-text TEXT  按文本匹配标题（支持子串匹配，多个用逗号分隔）
```

## 交互命令

运行过程中可用的交互命令：

| 命令 | 说明 |
|---|---|
| 回车 | 确认当前绑定正确，执行截图 |
| `r` | 重新 bind 当前活跃标签页（修正绑定窗口） |
| `s` | 跳过当前小节 |
| `q` | 退出并保存已完成的部分 |
