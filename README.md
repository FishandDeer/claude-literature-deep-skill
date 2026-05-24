# literature-deep-skill

> 一个为 [Claude Code](https://claude.com/claude-code) 设计的**学术文献深度处理 Skill**:支持「逐字摘抄、归纳总结、叙述改写、要素抽取、PPT 生成」五类原子能力,专门服务于**中文学术文献综述与汇报场景**。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Claude Code Skill](https://img.shields.io/badge/Claude%20Code-Skill-blueviolet)](https://docs.claude.com/en/docs/claude-code/skills)

---

## ✨ 为什么需要这个 Skill?

通用的 `paper-summarizer` 默认按"总结归纳"一种模式工作。但在真实的学术汇报场景中,用户的需求往往是**三种模式严格区分**的:

| 模式 | 用户原话 | 产出 |
|---|---|---|
| **MODE-A 逐字摘抄** | "直接复制""不要总结""我自己写 PPT 用" | 100% 原文,零改写零归纳,按文献逐篇组织 |
| **MODE-B 归纳总结** | "总结""归纳""帮我看看这篇讲了什么" | 结构化中文摘要(贡献+方法+实验+局限+点评) |
| **MODE-C 完整叙述** | "做成研究背景章节""不要只给论点,我要完整叙述" | 多篇文献融合的完整段落长文,可直接用作论文章节 |

**三种模式绝不能混淆**。这是本 Skill 第一原则——它源自多次"模式被偷偷混淆导致反复返工"的失败案例,详见 [`references/pitfalls.md`](references/pitfalls.md)。

此外,本 Skill 还内置了对中文学术 markdown 章节标题多种命名风格(`## 3.` / `## 三、` / `## 研究背景与动机` / `## 背景与动机`)的**容错抽取**——历史经验表明,仅依赖数字编号匹配会漏掉相当比例的文献。

---

## 🚀 安装

### 方法 1:直接 clone 到 Claude Code skills 目录

```bash
# Linux / macOS
git clone https://github.com/FishandDeer/claude-literature-deep-skill.git \
    ~/.claude/skills/literature-deep-skill

# Windows (PowerShell)
git clone https://github.com/FishandDeer/claude-literature-deep-skill.git `
    $env:USERPROFILE\.claude\skills\literature-deep-skill
```

### 方法 2:作为子模块或本地软链接

```bash
git clone https://github.com/FishandDeer/claude-literature-deep-skill.git
ln -s "$(pwd)/claude-literature-deep-skill" ~/.claude/skills/literature-deep-skill
```

### 依赖

- **Bash + awk**(`extract_sections.sh` 使用,Linux/macOS 自带)
- **Python ≥ 3.9 + `python-pptx`**(可选,仅生成 PPT 时需要)

  ```bash
  pip install python-pptx Pillow
  ```

- **中文字体**(Linux 端生成 PPT 时):

  ```bash
  sudo apt install fonts-noto-cjk
  ```

---

## 🎯 使用

在 Claude Code 中,直接用自然语言触发即可。Skill 会在动手前用 `AskUserQuestion` 与你确认模式:

```
> 帮我处理 /papers/ 下的若干篇文献,做成研究背景汇报
```

Claude 会问你:

1. **模式**:摘抄 / 总结 / 叙述 / 全套?
2. **要素**:研究背景 / 相关工作 / 数据集 / 方法 / 实验 / 全部?
3. **是否需要 PPT**?

确认后,Skill 会:

- MODE-A → 调用 `scripts/extract_sections.sh` 关键词 fuzzy match 抽取原文
- MODE-B → 按 `references/summary-templates.md` 模板生成 10 节结构化摘要
- MODE-C → 先 A 拿素材再以完整段落写成长文
- PPT 需求 → 调用 `scripts/build_pptx.py` 的 `AcademicDeck` 类(深蓝主色+橙色强调,16:9,Microsoft YaHei)

---

## 📂 仓库结构

```
literature-deep-skill/
├── SKILL.md                          # Skill 主入口:6 类能力 + 3 模式核心原则 + 5 步工作流
├── references/
│   ├── extraction-modes.md           # MODE-A/B/C 详细判别与执行细则
│   ├── section-mapping.md            # 中文学术 markdown 章节标题别名映射表
│   ├── summary-templates.md          # MODE-B 单篇 + 00-OVERVIEW 标准模板
│   ├── ppt-design.md                 # 学术 PPT 视觉规范(色彩/字体/版式/8 类布局)
│   └── pitfalls.md                   # 历史踩坑清单 + 交付前自检 checklist
├── scripts/
│   ├── extract_sections.sh           # 关键词 fuzzy match 抽取(single/batch 双模式)
│   └── build_pptx.py                 # 学术风 python-pptx 模板(AcademicDeck 类)
├── LICENSE                           # MIT
└── README.md
```

---

## 🧪 脚本独立使用(不通过 Claude)

两个核心脚本也可以脱离 Claude Code 直接使用。

### `extract_sections.sh`

```bash
# 1. 单文件单章节抽取(带 fallback 关键词)
./scripts/extract_sections.sh single paper.md "研究背景" "背景与动机"

# 2. 批量摘抄汇编
cat > papers.list <<EOF
paper-a.md|论文 A 的中文标题
paper-b.md|论文 B 的中文标题
EOF

./scripts/extract_sections.sh batch papers.list out.md \
    "研究背景:背景与动机" "相关工作"
```

### `build_pptx.py`

```python
from build_pptx import AcademicDeck

deck = AcademicDeck()
deck.add_cover(title="研究汇报", subtitle="组会报告",
               author="张三", date="2026-05")
deck.add_toc(["研究背景", "相关工作", "方法", "实验"])
deck.add_section_header("01", "研究背景", "问题定义与现状")
deck.add_content_slide(
    title="行动式标题示例:核心论点放在标题里",
    bullets=[
        ("第一个论据,带数据 [#1]", None),
        ("第二个论据,带次级注释 [#2]", "次级注释用更小字号、灰色"),
    ],
)
deck.save("output.pptx")
```

CLI 自检:

```bash
python3 scripts/build_pptx.py /tmp/demo.pptx
```

---

## 📜 设计哲学:三种模式

### MODE-A 逐字摘抄

把源文献中的目标章节,**100% 原封不动**地复制到输出,零改写、零归纳、零删减、零重组。**必须用脚本抽取**——人手处理一定会偷偷归纳。**必须按论文逐篇组织**——绝不按 PPT 页/主题/范式重组。

### MODE-B 归纳总结

10 节模板:基本信息表 / TL;DR / 研究背景 / 核心贡献 / 方法论 / 实验结果 / 局限性 / 个人点评 / 关键词 / 相关工作。每篇文献一个独立 `.md`,批量时另出 `00-OVERVIEW.md` 总览索引。中文优先,英文术语保留。

### MODE-C 完整叙述

每节由 2~5 个完整段落组成,而非 bullet list。数据嵌入文中(如「某权威机构 X 年报告显示某地区当年发生 N 起事故[#1]」),而非「• N 例 [#1]」。学术语气,克制,避免营销话术。

详见 [`references/extraction-modes.md`](references/extraction-modes.md)。

---

## 🤝 贡献

欢迎 PR。新增失败案例时,**请直接追加到 [`references/pitfalls.md`](references/pitfalls.md)** 的对应小节,这是本 Skill 不再重复踩坑的唯一机制。

格式:

```markdown
### 案例 X.Y:[一句话症状]
- **现象**:[做了什么,结果是什么]
- **退回反馈**:"[用户原话,务必逐字]"
- **根因**:[为什么会这样]
- **修复**:[如何改的]
- **教训**:[下次怎么避免,要可执行]
```

---

## 📄 License

[MIT](LICENSE)

---

## 🙏 致谢

本 Skill 的设计源自真实学术综述工作流中**多次反复返工**的经验沉淀——所有踩坑都被忠实记录在 [`references/pitfalls.md`](references/pitfalls.md) 中,这是它**最值钱的部分**。

如果它对你的学术工作有帮助,欢迎 star ⭐。也欢迎提 PR 把你自己遇到的失败案例追加到 pitfalls 里。
