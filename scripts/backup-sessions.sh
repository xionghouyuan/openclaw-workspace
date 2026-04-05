#!/bin/bash
# Backup OpenClaw session transcripts to GitHub
# Sanitizes tokens and secrets before committing

WORKSPACE="$HOME/.openclaw/workspace"
SESSION_DIR="$HOME/.openclaw/agents/main/sessions"
BACKUP_DIR="$WORKSPACE/sessions"

mkdir -p "$BACKUP_DIR"

# Sanitize and copy session .jsonl files
for f in "$SESSION_DIR"/*.jsonl; do
    [ -f "$f" ] || continue
    filename=$(basename "$f")
    # Remove tokens, passwords, API keys
    # Also remove user's sudo password from commands/outputs
    sed -E \
        -e 's/ghp_[A-Za-z0-9]{20,}/[GITHUB_TOKEN]/g' \
        -e 's/xhy0804\./[PASSWORD_REDACTED]/g' \
        -e 's/"token":\s*"[^"]*"/"token": "[REDACTED]"/g' \
        -e 's/sk-[A-Za-z0-9]{20,}/[API_KEY]/g' \
        "$f" > "$BACKUP_DIR/$filename"
done

cd "$WORKSPACE" || exit 1

# Check if there are any new or changed files
if [ -z "$(git status --porcelain sessions/)" ]; then
    echo "No session changes to backup."
    exit 0
fi

# Commit with timestamp
git add sessions/
git commit -m "Session backup - $(date '+%Y-%m-%d %H:%M:%S')"

# Push
git push origin master 2>&1

echo "Session backup done: $(date)"
