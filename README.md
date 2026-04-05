# 🚢 Freight Agent — 一人公司 AI Agent 团队

> Brandon 的 Freight Forwarder 创业工具箱
> 4 个 AI Agent，覆盖从获客到收款的全流程

---

## 📁 文件结构

```
freight-agent/
├── sales-support/          💼 Sales Support Agent（销售支持）
├── customer-service/       🎧 Customer Service Agent（客户服务）
├── operations/             ⚙️ Operations Agent（运营操作）
└── accounts/               💰 Accounts Agent（财务会计）
```

---

## 💼 Sales Support Agent — 销售支持

> 负责：获客、报价、跟进、社媒

| 文件 | 功能 |
|------|------|
| `Quote_Generator.xlsx` | ⭐ 报价单生成器 — 填黄色格子，自动出专业报价单 |
| `Margin_Calculator.xlsx` | 利润计算器 — 单票/多票 margin 自动算 + 速查表 |
| `Customer_CRM.xlsx` | 客户 CRM — 客户总览/报价记录/跟进提醒/统计看板 |
| `Followup_Agent.xlsx` | 跟进管理 — 今日任务/记录/客户分级VIP-A-B/消息模板 |
| `WhatsApp_Templates.xlsx` | WhatsApp/微信话术库 — 14个场景快速回复 |
| `Social_Media_Agent.xlsx` | 社媒内容 — LinkedIn/FB模板/内容日历/素材库 |

### 工作流程
```
客户询价 → 查运价(Rate DB) → 填报价单 → 发给客户 → 设跟进提醒 → 社媒持续曝光
```

---

## 🎧 Customer Service Agent — 客户服务

> 负责：客服、货物追踪、自动回复

| 文件 | 功能 |
|------|------|
| `Customer_Service_Agent.xlsx` | 客服工具箱 — FAQ中英双语/货物追踪/工单管理/客服话术 |
| `whatsapp-bot/` | WhatsApp 自动回复机器人 — 关键词识别 + 自动回复 |

### whatsapp-bot 文件说明
| 文件 | 用途 |
|------|------|
| `bot.py` | 主程序 |
| `config.json` | 营业时间 + 关键词配置 |
| `replies.json` | 回复模板（中英双语） |
| `dashboard.html` | Web 管理面板 |

### 支持的自动回复
| 触发关键词 | 自动回复 |
|-----------|---------|
| price / 报价 / 多少钱 | 引导客户提供信息以便报价 |
| where / 到哪 / 状态 | 要订单号帮你查货物状态 |
| time / 多久 / 时效 | 给出海运/空运参考时效 |
| service / 服务 / 能运 | 列出所有服务项目 |
| hi / hello / 你好 | 问候 + 引导 |

---

## ⚙️ Operations Agent — 运营操作

> 负责：文档、运价、航线、船期

| 文件 | 功能 |
|------|------|
| `Document_Agent.xlsx` | 文档生成 — 货物信息→发票/装箱单/文档清单/船期跟踪 |
| `Rate_Database.xlsx` | 运价数据库 — FCL/LCL/Air/Surcharges |
| `Rate_Trend_Tracker.xlsx` | 运价趋势 — 每周记录+变化对比+趋势看板 |
| `East_MY_Calculator.xlsx` | 东马转运 — PKL→Kuching/KK费用拆解 |
| `Quotation_Template.xlsx` | 备用报价单模板 |

### 工作流程
```
确认出货 → 填货物信息 → 自动生成发票+装箱单 → 跟踪船期 → 更新运价趋势
```

---

## 💰 Accounts Agent — 财务会计

> 负责：收支、应收应付、财务报表

| 文件 | 功能 |
|------|------|
| `Finance_Agent.xlsx` | 财务管理 — 收支记录/应收账款AR/应付账款AP/月度报表/财务看板 |

### 核心功能
- 💵 每笔收支记录（海运费收入、运费支出、报关费等）
- 📥 应收账款：谁欠你钱 + 到期提醒
- 📤 应付账款：你欠谁钱 + 付款日提醒
- 📊 月度报表：每月收入/支出/利润
- 📈 财务看板：总收入、总支出、净利润、利润率、逾期预警

---

## 🚀 快速开始

### 1. 首次设置
- 打开 `Quote_Generator.xlsx` → 填上你的公司信息
- 打开 `Rate_Database.xlsx` → 找供应商要运价填进去
- 打开 `Customer_CRM.xlsx` → 填上你的公司信息

### 2. 日常使用
| 场景 | 用哪个 |
|------|--------|
| 客户问价 | `sales-support/Quote_Generator.xlsx` |
| 查报价记录 | `sales-support/Customer_CRM.xlsx` → 报价记录 |
| 今天该联系谁 | `sales-support/Followup_Agent.xlsx` → 今日跟进 |
| 生成发票/装箱单 | `operations/Document_Agent.xlsx` |
| 跟踪货物 | `customer-service/Customer_Service_Agent.xlsx` → Tracker |
| 记收支 | `accounts/Finance_Agent.xlsx` |
| 发 LinkedIn | `sales-support/Social_Media_Agent.xlsx` → LinkedIn 模板 |
| 快速回复 WhatsApp | `sales-support/WhatsApp_Templates.xlsx` |

### 3. 每周检查
- [ ] 更新 `operations/Rate_Database.xlsx` 运价
- [ ] 记录 `operations/Rate_Trend_Tracker.xlsx` 本周运价
- [ ] 检查 `sales-support/Followup_Agent.xlsx` 待跟进客户
- [ ] 查看 `accounts/Finance_Agent.xlsx` 应收账款
- [ ] 发 2-3 条 LinkedIn/Facebook（`sales-support/Social_Media_Agent.xlsx`）

---

## 📊 Agent 架构图

```
                    ┌─────────────────┐
                    │   客户 Customer  │
                    └────────┬────────┘
                             │
              ┌──────────────┼──────────────┐
              │              │              │
    ┌─────────▼───┐  ┌──────▼──────┐  ┌───▼─────────┐
    │ Sales Support│  │Customer Svc │  │  Social Media│
    │  报价+跟进   │  │ 客服+追踪   │  │  LinkedIn/FB │
    └──────┬──────┘  └──────┬──────┘  └─────────────┘
           │                │
    ┌──────▼────────────────▼──────┐
    │       Operations 运营         │
    │  文档/运价/船期/东马转运       │
    └──────────────┬───────────────┘
                   │
    ┌──────────────▼───────────────┐
    │        Accounts 财务          │
    │   收支/应收应付/月度报表       │
    └──────────────────────────────┘
```

---

## 🔧 开发者说明

所有 Excel 文件由 Python + openpyxl 生成，源码在 `create_*.py` 文件中。

重新生成所有工具：
```bash
python3 create_rate_database.py
python3 create_generator.py
python3 create_crm.py
python3 create_templates_and_calc.py
python3 create_document_agent.py
python3 create_followup_agent.py
python3 create_finance_agent.py
python3 create_social_agent.py
python3 create_cs_agent.py
python3 create_upgrades.py
```
