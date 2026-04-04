# TOOLS.md - Local Notes

Skills define _how_ tools work. This file is for _your_ specifics — the stuff that's unique to your setup.

## What Goes Here

Things like:

- Camera names and locations
- SSH hosts and aliases
- Preferred voices for TTS
- Speaker/room names
- Device nicknames
- Anything environment-specific

## Examples

```markdown
### Cameras

- living-room → Main area, 180° wide angle
- front-door → Entrance, motion-triggered

### SSH

- home-server → 192.168.1.100, user: admin

### TTS

- Preferred voice: "Nova" (warm, slightly British)
- Default speaker: Kitchen HomePod

### Network Proxy (Mihomo)

- Config: `~/.config/mihomo/config.yaml`
- Binary: `/usr/local/bin/mihomo` (v1.19.22)
- Mixed port: `127.0.0.1:7890` (HTTP/SOCKS5 proxy)
- REST API: `http://127.0.0.1:9090`
- Subscription: `https://ccsub.pz.pe/subscribe/Q4XLI2SI5ZN17SNJ`
- Startup: `nohup mihomo -d ~/.config/mihomo/ > ~/.config/mihomo/mihomo.log 2>&1 &`
- Enable at login: add to startup applications or systemd
```

## Why Separate?

Skills are shared. Your setup is yours. Keeping them apart means you can update skills without losing your notes, and share skills without leaking your infrastructure.

---

Add whatever helps you do your job. This is your cheat sheet.
