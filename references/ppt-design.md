# 学术 PPT 视觉规范与版式(V6 风格)

> 适用场景:中文学术汇报、组会、论文阅读会、开题/中期/答辩。**英文国际会议 PPT 不在本规范范围内**。

本规范由 `scripts/build_pptx.py` 中的 `AcademicDeck` 类完整实现。**生成 PPT 时优先调用类方法,不要从 `Presentation()` 裸写**。

---

## 一、基础参数

| 参数 | 值 |
|---|---|
| 比例 | 16:9(`Inches(13.333, 7.5)`) |
| 中文字体 | Microsoft YaHei(首选)→ Noto Sans CJK SC(Linux fallback)|
| 英文/数字字体 | Calibri(与中文同行混排不至于错位)|
| 大标题字号 | 40 pt(封面)、48 pt(PART 扉页)|
| 内容页主标题 | 22 pt 加粗(放在富 header 中) |
| Eyebrow | 10 pt 加粗、橙色(放在标题上方)|
| Subtitle | 12 pt 灰色斜体(放在标题下方,可选)|
| 正文字号 | 18~20 pt(主论点)、13~14 pt(次级注释)|
| 表格字号 | 表头 12 pt,正文 11 pt |
| 卡片正文 | 11 pt + 1.20 倍行距 |
| 注释/页脚 | 9~11 pt |

---

## 二、配色(V6 调色板)

**深蓝主色 + 橙色强调 + 5 种 Tint 浅底色**,经多次组会反馈验证为"学术、克制、不土"的组合。

### 主色系

| 用途 | RGB | 十六进制 | 常量名 |
|---|---|---|---|
| 主色(标题、章节条) | (31, 58, 104) | `#1F3A68` | `PRIMARY` |
| 副色(渐变、对比首列) | (46, 90, 140) | `#2E5A8C` | `PRIMARY_2` / `SECONDARY` |
| 强调色(eyebrow、引用编号、关键数据) | (197, 90, 17) | `#C55A11` | `ACCENT` |
| 学术绿(对比第三列、肯定信号) | (46, 125, 50) | `#2E7D32` | `GREEN` |
| 学术金(对比警示、第四模块) | (201, 154, 42) | `#C99A2A` | `GOLD` |
| 文字正色 | (34, 34, 34) | `#222222` | `DARK` |
| 文字次色 | (107, 107, 107) | `#6B6B6B` | `MUTED` |
| 浅灰背景 | (230, 233, 238) | `#E6E9EE` | `GRAY_L` |
| 中灰分隔线 | (201, 207, 215) | `#C9CFD7` | `GRAY_M` |

### Tint 浅底色(用于卡片/对比页)

| 主题 | RGB | 十六进制 | 常量名 |
|---|---|---|---|
| 蓝(中性、首选) | (235, 241, 248) | `#EBF1F8` | `TINT_BLUE` |
| 橙(强调、关键发现) | (251, 236, 221) | `#FBECDD` | `TINT_ORG` |
| 绿(肯定、达成共识) | (231, 241, 232) | `#E7F1E8` | `TINT_GRN` |
| 红(警示、空白领域) | (251, 230, 230) | `#FBE6E6` | `TINT_RED` |
| 金(进度、待办) | (255, 246, 224) | `#FFF6E0` | `TINT_GOLD` |

**禁止**:
- 深色背景(看着累、投影仪偏色严重)— **PART 扉页/结尾页除外**,这两类用满版深蓝是规范允许的。
- 同一张 slide 出现 > 5 种颜色(评委会觉得花)。
- 红色用于"强调正向数据"(易与"错误"语义混淆)— 红只用于警示、空白、不足。

---

## 三、版式模板(11 类,均由 AcademicDeck 提供)

| 版式 | 方法 | 关键元素 |
|---|---|---|
| **封面页** | `add_cover` | 左侧深蓝竖条 + 橙色细竖条 + eyebrow + 大标题(40 pt)+ 橙色短分隔条 + 副标题 + 元信息(汇报人/单位/日期)|
| **目录页** | `add_toc` | 富 header + 4 行罗马编号块 + 章节名 + tagline(支持简单字符串列表回退)|
| **PART 扉页** | `add_part_divider` | 满版深蓝 + 大号 PART 编号(110 pt 橙)+ 章节名(48 pt 白)+ 橙色短线 + tagline(灰斜体)|
| **内容页(标准)** | `add_content_slide` | 富 header(eyebrow + title + subtitle + 分隔线)+ bullets(主+次级注释)|
| **对比页** | `add_compare_slide` | 富 header + 2~3 列等宽,每列一种主题色(蓝/橙/绿)|
| **卡片网格页** | `add_card_grid` | 富 header + 2~6 张 Tint 卡片,每张卡片左侧色条 + 标题 + bullets,支持 5 种主题 |
| **表格页** | `add_table_slide` | 富 header + 深蓝表头 + 隔行 GRAY_L 底色 + accent_rows 橙色高亮 + 可选 caption |
| **图表强调页** | `add_figure_slide` | 富 header + 居中大图(等比缩放)+ 底部一句话结论 |
| **参考文献页** | `add_references_slide` | 富 header + 12 pt 编号列表(编号橙色加粗)|
| **结尾页** | `add_ending` | 满版深蓝 + 大字"谢谢/Q&A"(72 pt)+ 橙色短线 + 联系方式 + 可选 sub_text |
| **演讲者备注** | `AcademicDeck.add_notes(slide, text)` | 给任意 slide 添加 notes(自动套用 YaHei 11 pt)|

---

## 四、富 Header 标准

每张内容页(非 cover、非 part_divider、非 ending)都使用统一的 4 层 header:

```
┌─────────────────────────────── 顶部深蓝条(0.10")──────────────┐
├─────────────────────────────── 橙色细条(0.04")────────────────┤
│  [eyebrow 橙色 10pt]                              [NN / TOTAL] │
│  [主标题 22pt PRIMARY 加粗]                                     │
│  [副标题 12pt MUTED 斜体(可选)]                                │
├──────────────── GRAY_M 灰色分隔线 ──────────────────────────────┤
```

- **eyebrow**:章节定位,如 ``"Ⅱ · 研究现状"``、``"Method · Section 3.2"``
- **title**:行动式标题,放结论(见第五节)
- **subtitle**(可选):一句话支撑论据或数据点
- **页码**:右上 ``NN / TOTAL`` 格式(若 `total_pages` 未指定则只显示 `NN`)

---

## 五、Action Title(行动式标题)

每张内容页的标题必须传达**结论**,而不是**主题**。

| ❌ 主题式 | ✅ 行动式 |
|---|---|
| "数据集" | "在 4 个公开数据集上验证,规模 2k~50k" |
| "实验结果" | "Top-1 准确率从 89.2% 提升至 94.7%" |
| "相关工作" | "现有方法在长尾分布下精度下降 12 pp" |
| "方法" | "通过 CBAM 双注意力实现跨模态对齐" |
| "VLM 研究现状" | "VLM 在 DMS 形成三大范式,2024 年是范式革命之年" |

---

## 六、内容密度

- **每页一个核心论点**。超过两个就分页。
- **正文 ≤ 6 行 / 页**;每行 ≤ 35 个中文字符或 70 个英文字符。
- **图 > 表 > 文字**:能用图就别用表,能用表就别用文字段落。
- **引用编号 `[#N]`** 用橙色,放在被引论点末尾(在 bullet 文本内嵌入即可)。
- **卡片网格**:不要超过 6 张,文字超过 3 行就拆成两页。

---

## 七、必须避免的"翻车现场"

| 现象 | 根因 | 对策 |
|---|---|---|
| 中文字符显示成方框 | Linux 端无 YaHei | 用 Noto Sans CJK SC,或预先 `apt install fonts-noto-cjk` |
| 英文/数字混排乱码 | 没有显式设置 `eastAsia` typeface | `_set_run` 自动注入 `a:eastAsia` 字体节,无需额外处理 |
| 文字溢出文本框 | 没估算字符数 | 标题 ≤ 16 个中文字符;超过的拆成 title + subtitle |
| 图片被拉变形 | 修改 width 没同步 height | `add_figure_slide` 内部用 PIL 读原图比例等比缩放 |
| PART 扉页大字号挤掉章节名 | 长章节名 + 110pt 编号 | 章节名 ≤ 8 个汉字;tagline 单独一行 |
| 上下页字体大小忽变 | 多人协作 / 直接 set placeholder | 全程走 `AcademicDeck`,字号在常量层统一 |
| 配图分辨率糊 | 截图低 dpi | `pdftoppm -r 200`,自绘图 mermaid → png 300 dpi |
| 引用编号颜色不一 | 老脚本里手填 RGB | 在 bullet 文本里写 `[#1]` 即可,后续可统一染色 |

---

## 八、推荐架构

### 8.1 paper-reading PPT(单篇文献,15~20 min · ~12 页)

```
1.  add_cover                        # 论文元信息
2.  add_toc(["背景", "方法", "实验", "贡献", "启示"])
3.  add_part_divider("Ⅰ", "研究背景与动机")
4-5. add_content_slide × 2           # 背景 + 现有方法不足
6.  add_part_divider("Ⅱ", "方法")
7.  add_figure_slide                 # 总览图
8-9. add_content_slide × 2           # 模块细节
10. add_table_slide                  # 主结果对比
11. add_card_grid                    # 贡献 + 局限 + 对我们的启示
12. add_ending("Q & A")
```

### 8.2 综述报告 PPT(批量文献,25~40 min · ~30 页)

```
1.   add_cover
2.   add_toc(rich, 4 大模块)
3.   add_part_divider("Ⅰ", "研究背景")
4-9. add_content_slide × 6           # 时代窗口 / 法规 / 标准 / 现实 / 张力
10.  add_part_divider("Ⅱ", "研究现状")
11.  add_content_slide               # 阵营划分总览
12-15. add_card_grid × N             # 每阵营一张卡片网格
16.  add_table_slide                 # 跨论文对比大表
17.  add_compare_slide               # 共识 vs 空白
18.  add_part_divider("Ⅲ", "研究问题切入")
19-20. add_content_slide × 2
21.  add_references_slide
22.  add_ending
```

---

## 九、与 `scripts/build_pptx.py` 的关系

`scripts/build_pptx.py` 把以上规范固化为函数。生成 PPT 时:

1. **优先调用 `AcademicDeck` 类方法**;
2. 类方法不够用时,先在本文件补规范,再去脚本里补函数,**禁止反向**(先写代码再补文档);
3. 颜色常量改动只改 `build_pptx.py` 顶部的 V6 调色板段,全局生效;
4. CLI 自检:`python3 build_pptx.py /tmp/demo.pptx` 一键生成 9 页 demo,涵盖所有版式。

---

## 十、API 速查

```python
from build_pptx import AcademicDeck

d = AcademicDeck(total_pages=30, footer_text="组会 · 2026-05")

# 封面
d.add_cover(title=..., subtitle=..., eyebrow=..., author=..., date=..., org=..., meeting_tag=...)

# 目录(富格式三元组 / 简单字符串两种回退)
d.add_toc([("Ⅰ", "研究背景", "tagline 一句话"), ...])

# PART 扉页
d.add_part_divider("Ⅰ", "研究背景", tagline="...", total_parts=4)

# 内容页(eyebrow 在标题上方,subtitle 在标题下方)
d.add_content_slide(title=..., eyebrow=..., subtitle=...,
                    bullets=[(主, 次), ...])

# 对比页(2~3 列,自动套蓝/橙/绿主题)
d.add_compare_slide(title=..., columns=[(列名, [bullets]), ...])

# 卡片网格(2~6 张,主题:blue / orange / green / red / gold)
d.add_card_grid(title=..., cards=[(卡片名, [bullets], "blue"), ...])

# 表格(第一行表头,accent_rows 高亮)
d.add_table_slide(title=..., data=[[...], ...],
                  accent_rows=[2], col_widths=[2, 3, 1, 1.5],
                  caption="表 1 ...")

# 图表强调
d.add_figure_slide(title=..., image_path=..., caption=...)

# 参考文献
d.add_references_slide(refs=[...])

# 结尾
d.add_ending("Q & A", contact="...", sub_text="...")

# 备注(给任意已 add 的 slide)
d.add_notes(slide, "演讲者备注内容")

d.save("/path/to/output.pptx")
```
