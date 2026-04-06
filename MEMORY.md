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

## PocketClaw 图片发送（2026-04-05）

### 问题
在 PocketClaw 对话中无法发送本地图片给对方。

### 错误判断
我说「PocketClaw 不支持显示图片」，是错的。实际上：
- PocketClaw 支持图片，但需要通过 `clawpilot send` 命令发送
- 不能通过文件路径引用，必须用 `clawpilot send <绝对路径>`
- 图片会先上传到 OSS（pocketclaw.oss-cn-hongkong.aliyuncs.com），再发到会话

### 解决
- clawpilot 版本需更新到最新：`npm install -g @rethinkingstudio/clawpilot@latest --prefix ~/.npm-global`
- skill 已安装：`npx skills add Rethinking-studio/clawpilot-skills --skill clawpilot-send --global`
- 使用 `clawpilot send "/path/to/file"` 发送

### 教训
**之前已经看到过 `clawpilot` 命令，但没用它。** 当问题出现时，应该先检查已有的工具和技能，而不是从头查文档。

---

## 反思：为什么之前学不会

**直接原因：** 遇到问题就去查文档，没有先盘点已有的工具和能力。

**根因：** 
1. **工具清单没有内化。** 桌面上明明有 clawpilot，却没有用它来解决问题。
2. **文档路径依赖。** 遇到新问题先想「我要查文档」，而不是「我要看看手边有什么」。
3. **过早下结论。** 「PocketClaw 不支持图片」——这是基于不完整信息的判断。
4. **没有追问用户是否已有解决方案。** 用户提到 PocketClaw，却没有问「这个工具怎么发图片」。

**以后遇到类似问题的正确姿势：**
1. 先问用户「这个问题你之前怎么解决的？」
2. 检查已有的 skills、workspace 工具、npm 全局命令
3. 不确定时先说「我查一下」而不是直接下结论
4. 用户给 GitHub 链接 → 直接安装并使用，而不是先研究原理

---

## Reddit 内容扫描与整理 (2026-04-06)

### 经验教训
**问题：** 最开始找的内容质量不高，被熊老大批评「越来越水」。

**原因：** 
1. Scanner 只按 upvotes/评论数筛选，没有评估内容深度
2. 扫到的大多是链接分享、标题党、简单新闻
3. 真正有深度的是需要点进去读正文+评论区精选的

**解决方案：**
1. 手动整理 > 自动化扫描（自动化只能做初筛）
2. 严格标准：self-post 300+字+30+评论，或外链 2000+ upvotes+100+评论
3. 板块选择：truereddit、philosophy、DepthHub、skeptic、askscience、lectures
4. 排除：politics、technology（容易水）、生活琐事、个人故事

### 工具使用
- `clawpilot send <file>` 发送文件给用户
- Reddit API 通过 mihomo 代理访问：`curl -x http://127.0.0.1:7890`
- Reddit 用户代理要够真实否则被拦

---

## openclaw browser 工具 (2026-04-06)

### 问题
browser 工具在 Linux VM 上不稳定，经常报 "tab not found" 错误。

### 原因
Linux VM 有 DISPLAY=:0（图形服务器），但 Chrome 需要额外参数才能在无头模式运行。

### 解决
在 openclaw.json 中添加：
```json
"browser": {
  "noSandbox": true
}
```

### 可用命令
- `openclaw browser start` / `stop`
- `openclaw browser open <url>`
- `openclaw browser navigate <url>`
- `openclaw browser screenshot [path]`
- `openclaw browser snapshot [--limit N]`
- `openclaw browser click <ref>`
- `openclaw browser press <key>` (PageDown, etc.)
- `openclaw browser tabs`
- `openclaw browser status`

### 限制
Reddit 有 CAPTCHA 检测，自动化浏览器访问会被拦截。API 访问可行。

---

## Reddit Scanner 脚本 (2026-04-06)

### 位置
`~/.openclaw/workspace/scripts/reddit-scanner-v2.py`

### 功能
- 每3小时扫描指定板块
- 过滤低质量内容
- 用缓存避免重复推送
- 通过 clawpilot send 自动发给用户

### Cron Job
- Job ID: 5a2d2cc4-960c-4d81-941b-39d0a8f61f61
- schedule: `0 */3 * * *` (Asia/Shanghai)

### 缓存文件
`~/.openclaw/workspace/scripts/.reddit_seen_cache_v2`
