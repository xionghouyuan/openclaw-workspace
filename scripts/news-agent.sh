#!/bin/bash
# 新闻 AI 工作流：抓取 → 保存 → 生成 AI 分析素材
# 使用方法: bash news-agent.sh

export PATH="$HOME/.npm-global/bin:$PATH"
DATE=$(date +%Y%m%d_%H%M)
OUTDIR="$HOME/openclaw-news/$DATE"
REPORT="$OUTDIR/AI-BRIEFING.md"
mkdir -p "$OUTDIR"

log() { echo "[$(date +%H:%M:%S)] $1"; }

log "新闻抓取开始"

# 1. HackerNews Top 10
log "[1/5] HackerNews..."
opencli hackernews top --limit 10 -f json 2>/dev/null > "$OUTDIR/hackernews.json"
HN_COUNT=$(python3 -c "import json; print(len(json.load(open('$OUTDIR/hackernews.json'))))" 2>/dev/null || echo "0")
log "  → $HN_COUNT 条"

# 2. 36氪
log "[2/5] 36氪..."
opencli 36kr news --limit 10 -f json 2>/dev/null > "$OUTDIR/36kr.json"
log "  → $(python3 -c "import json; print(len(json.load(open('$OUTDIR/36kr.json'))))" 2>/dev/null || echo "0") 条"

# 3. BBC News
log "[3/5] BBC..."
opencli bbc news --limit 10 -f json 2>/dev/null > "$OUTDIR/bbc.json" && log "  → OK" || log "  ⚠ failed (跳过)"

# 4. ProductHunt
log "[4/5] ProductHunt..."
opencli producthunt hot --limit 10 -f json 2>/dev/null > "$OUTDIR/producthunt.json" && log "  → OK" || log "  ⚠ failed (跳过)"

# 5. Reddit
log "[5/5] Reddit hot..."
opencli reddit hot --limit 10 -f json 2>/dev/null > "$OUTDIR/reddit.json" && log "  → OK" || log "  ⚠ failed (跳过)"

log "抓取完成，汇总:"
ls -lh "$OUTDIR"

# 生成 AI 分析模板（供我下一步分析用）
python3 << 'PYEOF' > "$OUTDIR/RAW-DATA.txt"
import json, os

print("=" * 60)
print("HACKERNEWS TOP 10")
print("=" * 60)
try:
    with open(os.path.expanduser('$OUTDIR/hackernews.json')) as f:
        for item in json.load(f):
            print(f"[{item['rank']}] {item['title']}")
            print(f"    ★{item['score']} | 💬{item['comments']} | by {item['author']}")
            print(f"    {item['url']}\n")
except: pass

print("=" * 60)
print("36KR TOP 10")
print("=" * 60)
try:
    with open(os.path.expanduser('$OUTDIR/36kr.json')) as f:
        for item in json.load(f):
            print(f"[{item['rank']}] {item['title']}")
            print(f"    {item.get('date','')}\n")
except: pass

PYEOF

echo ""
echo "✅ 完成！数据在: $OUTDIR"
echo "📄 汇总文本: $OUTDIR/RAW-DATA.txt"
