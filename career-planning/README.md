# Career Planning — WorkBuddy 技能包

职业规划与简历生成技能，适用于求职规划、简历制作、岗位推荐等场景。

## 功能

1. **结构化信息收集** — 通过 15 题渐进式访谈，全面了解用户经验、能力与求职意向
2. **简历生成** — 自动生成专业排版的 DOCX 简历初稿（可下载、可编辑）
3. **岗位推荐** — 基于求职画像，搜索并推荐真实在招岗位

## 安装

将整个 `career-planning/` 目录放入 WorkBuddy 的 skills 目录：

- **用户级**（推荐）：`~/.workbuddy/skills/career-planning/`
- **项目级**：`<project>/.workbuddy/skills/career-planning/`

## 触发方式

在 WorkBuddy 对话中说出以下任意关键词即可触发：

- 找新工作、换工作、求职、简历
- 职业规划、岗位推荐
- 帮我写简历、改简历

## 文件结构

```
career-planning/
├── SKILL.md                    # 完整工作流说明
├── scripts/
│   └── generate_resume.py      # DOCX 简历生成脚本
└── README.md                   # 本文件
```

## 依赖

- Python 3.8+
- python-docx（`pip install python-docx`）

## 许可证

MIT
