#!/bin/bash
# Mihomo health check and auto-restart script
# Called by cron every 5 minutes

LOG="/home/xionghouyuan2/.openclaw/workspace/scripts/mihomo-health.log"
LOCK="/tmp/mihomo-health.lock"

# Prevent concurrent runs
exec 200>"$LOCK"
if ! flock -n 200; then
    exit 0
fi

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" >> "$LOG"
}

# Check if mihomo process is running
if ! pgrep -x "mihomo" > /dev/null 2>&1; then
    log "Mihomo not running, starting..."
    nohup mihomo -d /home/xionghouyuan2/.config/mihomo/ >> /home/xionghouyuan2/.config/mihomo/mihomo.log 2>&1 &
    sleep 3
    if pgrep -x "mihomo" > /dev/null 2>&1; then
        log "Mihomo started (pid $(pgrep -x mihomo))"
    else
        log "ERROR: Mihomo failed to start!"
    fi
    exit 0
fi

# Check if proxy actually works by testing github through it
HTTP_CODE=$(curl -s --connect-timeout 5 -o /dev/null -w "%{http_code}" -x http://127.0.0.1:7890 https://github.com 2>/dev/null)
if [ "$HTTP_CODE" != "200" ] && [ "$HTTP_CODE" != "301" ] && [ "$HTTP_CODE" != "302" ]; then
    log "Proxy not working (http $HTTP_CODE), restarting mihomo..."
    pkill -x mihomo
    sleep 2
    nohup mihomo -d /home/xionghouyuan2/.config/mihomo/ >> /home/xionghouyuan2/.config/mihomo/mihomo.log 2>&1 &
    sleep 3
    if pgrep -x "mihomo" > /dev/null 2>&1; then
        log "Mihomo restarted successfully (pid $(pgrep -x mihomo))"
    else
        log "ERROR: Mihomo failed to restart!"
    fi
else
    log "Mihomo OK (pid $(pgrep -x mihomo), proxy working)"
fi
