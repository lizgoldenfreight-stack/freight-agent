# 🤖 WhatsApp Auto-Reply Bot

自动回复 WhatsApp 消息的机器人。

## 文件说明

| 文件 | 用途 |
|------|------|
| `config.json` | 机器人配置（营业时间、关键词等） |
| `replies.json` | 自动回复模板（中英双语） |
| `bot.py` | 机器人主程序 |
| `dashboard.html` | 管理面板（浏览器打开） |

## 快速开始

```bash
# 1. 安装依赖
pip install pywhatkit flask

# 2. 打开 WhatsApp Web
# 浏览器访问 web.whatsapp.com

# 3. 运行机器人
python3 bot.py
```

## 支持的关键词

| 意图 | 英文关键词 | 中文关键词 |
|------|-----------|-----------|
| 报价 | price, how much, quote | 报价, 多少钱 |
| 追踪 | where, status, track | 到哪, 状态, 追踪 |
| 时效 | when, transit, time | 多久, 时效, 船期 |
| 服务 | service, can you | 服务, 能运 |
| 问候 | hi, hello, hey | 你好 |

## 自定义

- 编辑 `replies.json` 修改回复内容
- 编辑 `config.json` 修改营业时间和关键词
