# historiographer-skills

一个面向团队协作与开源共享的 AI skills 仓库，用于沉淀可复用的技能说明、辅助脚本与参考资料。

English version: [README.en.md](README.en.md)

## 仓库目标

这个仓库主要解决三类问题：

1. 把零散的个人 skill 沉淀成可复用资产
2. 让团队成员可以统一复用、维护和迭代 skill
3. 提供一个标准结构，方便通过 GitHub Fork / PR 方式持续贡献

这里的一个 skill，本质上就是一个独立目录，里面至少包含一个 `SKILL.md`，必要时再附带脚本、模板或参考资料。

## 仓库结构

```text
historiographer-skills/
  ├── README.md
  ├── README.en.md
  ├── CONTRIBUTING.md
  ├── LICENSE
  └── skills/
      ├── docx-add-screenshots/
      │   ├── SKILL.md
      │   └── scripts/
      │       └── docx_screenshot_inserter.py
      └── wsl-win-openbrowser/
          ├── SKILL.md
          └── references/
              └── .gitkeep
```

## Skill 目录约定

每个 skill 推荐保持自包含，目录内可以包含以下内容：

- `SKILL.md`
  核心说明文件，定义 skill 的用途、触发场景、流程、约束和输出
- `scripts/`
  辅助脚本，仅在 skill 需要真实执行逻辑时提供
- `references/`
  补充资料、命令模板、规范文档等
- `templates/`
  可复用模板文件
- `assets/`
  配图、示意资源或静态素材

最小可用结构是：

```text
skills/my-skill/
  └── SKILL.md
```

## 当前已收录 Skills

### `docx-add-screenshots`

用途：
为 Word 文档的指定章节批量插入浏览器截图，适合生成带图示的操作文档、交付文档或产品说明文档。

包含内容：

- `SKILL.md`
- `scripts/docx_screenshot_inserter.py`

### `wsl-win-openbrowser`

用途：
在 WSL 环境中通过 Windows 侧 `openCLI` 驱动浏览器，用于打开网页、等待加载、查看状态和截图。

包含内容：

- `SKILL.md`
- `references/`

## 如何使用这个仓库

常见用法有两种：

### 1. 直接作为技能仓库参考

你可以直接浏览 `skills/` 下的目录，查看每个 skill 的 `SKILL.md`，按需复制到你自己的 agent / CLI / 自动化环境中。

### 2. 作为团队内部共享仓库

团队成员把成熟的个人 skill 提交到本仓库，通过代码评审统一维护，逐步形成内部 skill 资产库。

## 如何贡献

1. Fork 本仓库
2. 新建分支
3. 在 `skills/<skill-name>/` 下新增或修改 skill
4. 确保至少包含一个 `SKILL.md`
5. 如果需要脚本或参考资料，只提交与该 skill 直接相关的内容
6. 提交 Pull Request

更详细的规范见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 编写 Skill 的建议

一个高质量的 skill，建议至少写清楚：

- 这个 skill 是解决什么问题的
- 什么场景下应该触发
- 需要哪些输入
- 执行流程是什么
- 有哪些限制或风险
- 最终输出什么结果

不要把 skill 写成笼统口号；应尽量让其他人拿到后可以直接复用。

## 命名与维护约定

- skill 目录名统一使用 `kebab-case`
- 一个目录只表达一个 skill
- 只保留与 skill 直接相关的内容
- 尽量避免把大量无关二进制文件提交进仓库
- 修改已有 skill 时，优先保持兼容，不做无关重构

## License

本仓库使用 [MIT License](LICENSE)。
