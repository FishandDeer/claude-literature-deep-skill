# 中文学术 markdown 章节标题别名映射

> 本表是脚本化抽取章节内容时的关键词字典。**所有抽取脚本必须用 fuzzy match,不能依赖固定数字编号**。否则会漏掉相当比例的文献(实际经验:仅匹配 `## 3`/`## 10` 阿拉伯数字编号时,中文数字与无编号风格的文献章节会全部为空)。

---

## 一、章节编号风格(同一份文件内通常一致,但不同文献之间会混用)

| 风格 | 示例 |
|---|---|
| 阿拉伯数字 + 点 | `## 1. 基本信息表` `## 3. 研究背景与动机` `## 10. 相关工作` |
| 中文数字 + 顿号 | `## 一、基本信息表` `## 三、研究背景(Background)` `## 十、相关工作(Related Work)` |
| 无编号 | `## 基本信息` `## 研究背景与动机` `## 相关工作` |
| 带英文括注 | `## 三、研究背景(Background)` `## 相关工作(Related Work)` |

---

## 二、目标章节别名表

### 研究背景

主关键词:`研究背景`
备用关键词:`背景与动机`、`背景`、`Background`、`Motivation`

完整别名清单:
```
研究背景
研究背景与动机
3. 研究背景
3. 研究背景与动机
三、研究背景
三、研究背景(Background)
背景与动机
背景与动机 (Background & Motivation)
Background
Motivation
```

### 相关工作

主关键词:`相关工作`
备用关键词:`Related Work`、`相关研究`、`相关工作推荐`

完整别名清单:
```
相关工作
10. 相关工作
十、相关工作
十、相关工作(Related Work)
相关工作(Related Work)
相关工作 (Related Work)
相关工作推荐
Related Work
相关研究
```

### 核心贡献

主关键词:`贡献`
完整别名清单:
```
核心贡献
关键贡献
主要贡献
4. 核心贡献
四、核心贡献
四、主要贡献(Key Contributions)
Key Contributions
Contributions
```

### 方法论

主关键词:`方法`
备用关键词:`Methodology`、`Method`
完整别名清单:
```
方法论
方法
5. 方法论
五、方法论
五、方法论(Methodology)
方法 (Methodology)
方法论:XXX 框架详解  # 带子标题的变体(注意冒号 + 自定义后缀)
Methodology
Approach
```

### 实验结果

主关键词:`实验`
备用关键词:`结果`、`Results`、`Experimental Results`
完整别名清单:
```
实验结果
实验
6. 实验结果
六、实验结果
六、实验结果(Results)
实验结果 (数据集 + 主要结果 + 消融实验)
Experimental Results
Results
Evaluation
```

### 数据集

数据集通常是**实验结果**节的子标题(`### 6.1 数据集` 或 `### 数据集`),也可能是**方法论**节的子标题。抽取时:

1. 先尝试匹配独立的 `## 数据集` 或 `### 数据集` 一级 / 二级标题
2. 否则在「实验结果」节内全文 grep `数据集` 关键词,抽取其下到下一个 `###` 的内容

### 局限性

主关键词:`局限`
完整别名清单:
```
局限性
局限与不足
局限性与未来工作
讨论与局限性
讨论与局限性(Limitations)
7. 局限性
七、局限性
七、讨论与局限性(Limitations)
七、局限与未来工作 (Limitations)
Limitations
```

### 个人点评

主关键词:`点评` 或 `评估`
完整别名清单:
```
个人点评
批判性评估
关键评估
关键评估 (Critical Assessment)
个人评价与批判性思考 (Critical Assessment)
Critical Assessment
Personal Comments
```

### 关键词

主关键词:`关键词`
完整别名清单:
```
关键词
关键词 / Tags
关键词(Keywords)
Keywords
Tags
```

---

## 三、推荐的 awk fuzzy match 模板

```bash
extract_section() {
  # $1 = source file, $2 = primary key, $3 = optional fallback key
  local f="$1" key="$2" key2="${3:-}"
  awk -v key="$key" -v key2="$key2" '
    /^## / {
      if (in_sec) { in_sec=0 }
      if (index($0, key) > 0) { in_sec=1; next }
      if (key2 != "" && index($0, key2) > 0) { in_sec=1; next }
    }
    in_sec { print }
  ' "$f"
}

# 用法:
extract_section "paper.md" "研究背景" "背景与动机" > bg.txt
extract_section "paper.md" "相关工作" "Related Work" > rw.txt
```

---

## 四、易踩坑案例

### 案例 1:缺少前缀字的章节漏抽

部分文献用的是 `## 背景与动机 (Background & Motivation)`,而不是 `## 研究背景`(没有"研究"前缀)。如果只匹配关键词 `研究背景`,这类文献会出空白。
**修复**:增加 `背景与动机` 作为 fallback,即调用 `extract_section "$src" "研究背景" "背景与动机"`。

### 案例 2:带子标题的章节漏抽

如果只匹配 `^## 5. 方法论$`(精确匹配),则 `## 5. 方法论:XXX 框架详解` 这种带冒号+自定义后缀的形式会漏。
**修复**:用 `index($0, "方法")` 包含式匹配,允许标题任意后缀。

### 案例 3:中文数字编号漏抽

如果只匹配 `## 3. 研究背景`(数字 + 点),则 `## 三、研究背景(Background)` 中文数字编号的版本会全部漏掉。
**修复**:始终用关键词文本匹配,完全无视编号。
