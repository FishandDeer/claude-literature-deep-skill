#!/bin/bash
# extract_sections.sh — 关键词 fuzzy match 的章节抽取脚本
#
# 用法 1(单文件单章节):
#   ./extract_sections.sh single <markdown_file> <key1> [<key2_fallback>]
#   示例: ./extract_sections.sh single paper.md "研究背景" "背景与动机"
#
# 用法 2(批量摘抄汇编):
#   ./extract_sections.sh batch <list_file> <out_file> [<key1>:<key2_fallback>]...
#   list_file 每行格式: <markdown_filename>|<论文标题>
#   示例:
#     ./extract_sections.sh batch papers.list out.md "研究背景:背景与动机" "相关工作"
#
# 原则(参见 references/pitfalls.md):
#   - 永远用关键词 fuzzy match,绝不用 ^## 3\. 这种数字正则
#   - 永远准备 fallback 关键词
#   - 输出按论文逐篇组织,不重组

set -e

extract_section() {
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

cmd_single() {
  local f="$1" key="$2" key2="${3:-}"
  if [[ ! -f "$f" ]]; then
    echo "ERROR: file not found: $f" >&2
    exit 1
  fi
  extract_section "$f" "$key" "$key2"
}

cmd_batch() {
  local list_file="$1"
  local out_file="$2"
  shift 2
  local sections=("$@")

  if [[ ! -f "$list_file" ]]; then
    echo "ERROR: list file not found: $list_file" >&2
    exit 1
  fi

  local dir
  dir="$(dirname "$(realpath "$list_file")")"

  : > "$out_file"
  {
    echo "# 文献章节逐字摘抄汇编"
    echo ""
    echo "> 本文档为**纯逐字摘抄**:章节内容均直接从源文献复制,不做改写、归纳、删减或重组。"
    echo "> 源目录:\`$dir\`"
    echo ""
    echo "---"
    echo ""
  } >> "$out_file"

  local idx=1
  while IFS='|' read -r fname title; do
    [[ -z "$fname" || "$fname" =~ ^# ]] && continue
    local src="$dir/$fname"
    if [[ ! -f "$src" ]]; then
      echo "WARN: source not found, skipping: $src" >&2
      continue
    fi

    {
      echo "## #$idx — $title"
      echo ""
      echo "**源文件**: \`$fname\`"
      echo ""
    } >> "$out_file"

    for sec in "${sections[@]}"; do
      local key="${sec%%:*}"
      local key2=""
      if [[ "$sec" == *:* ]]; then
        key2="${sec##*:}"
      fi

      {
        echo "### ${key}(原文逐字摘抄)"
        echo ""
      } >> "$out_file"

      local content
      content="$(extract_section "$src" "$key" "$key2")"
      if [[ -z "$content" ]]; then
        echo "WARN: empty section '$key' in $fname (tried fallback '$key2')" >&2
        echo "> ⚠️ 未在源文档中找到对应章节(已尝试关键词 \`$key\` 与 \`$key2\`)。" >> "$out_file"
      else
        echo "$content" >> "$out_file"
      fi
      echo "" >> "$out_file"
    done

    {
      echo "---"
      echo ""
    } >> "$out_file"

    idx=$((idx + 1))
  done < "$list_file"

  echo "Done. $(wc -l < "$out_file") lines written to $out_file" >&2
}

case "${1:-}" in
  single)
    shift
    cmd_single "$@"
    ;;
  batch)
    shift
    cmd_batch "$@"
    ;;
  *)
    cat >&2 <<'USAGE'
Usage:
  extract_sections.sh single <markdown_file> <key1> [<key2_fallback>]
  extract_sections.sh batch  <list_file> <out_file> <key1[:key2]>...

list_file format (one per line):
  filename.md|论文标题
  # 以 # 开头为注释行

Examples:
  extract_sections.sh single paper.md "研究背景" "背景与动机"
  extract_sections.sh batch papers.list out.md "研究背景:背景与动机" "相关工作"
USAGE
    exit 1
    ;;
esac
