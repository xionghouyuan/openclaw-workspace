# MEMORY.md - Long-Term Memory

## Who I Am
- Name: ClawBear 🐻
- Born: 2026-04-05
- Named by: 熊老大

## Who I'm Helping
- Name: 熊老大
- Timezone: Asia/Shanghai (GMT+8)
- SSH/git setup known

## Key Experience

### ClawPilot + PocketClaw Setup (2026-04-05)
熊老大在 Linux (VMware) 上配置 PocketClaw 连 OpenClaw gateway。

**Known issues:**
- `/usr/local` 全局 npm 安装需要 `--prefix ~/.npm-global`
- ClawPilot 必须用 `clawpilot set-token` 提供 gateway token，否则报 `unauthorized: gateway token missing`
- pm2 systemd 服务 `pm2-xionghouyuan2.service` 必须手动 `sudo systemctl start` 一次才能活
- `openclaw gateway restart` 会导致 ClawPilot 进程消失，但 pm2 会自动复活（如果 systemd pm2 服务在跑）

**Setup summary:**
- ClawPilot: 2.0.1, installed via npm to ~/.npm-global
- pm2 managed, systemd service: pm2-xionghouyuan2.service
- Gateway token saved in ~/.clawai/config.json
- pm2 save 已执行，进程列表在 ~/.pm2/dump.pm2

## Infrastructure
- GitHub user: xionghouyuan2
- GitHub token stored in ~/.git-credentials
- Workspace is a git repo (fresh, no remote yet)
