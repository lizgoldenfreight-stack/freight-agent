# 🚢 Freight Quotation Agent - 使用指南

## 文件清单

| 文件 | 用途 |
|------|------|
| `Rate_Database.xlsx` | 运价数据库 — 你从供应商拿到的报价全部填在这里 |
| `Quotation_Template.xlsx` | 报价单模板 — 复制一份，填上价格，发给客户 |
| `Customer_CRM.xlsx` | 客户管理 — 客户信息、报价记录、跟进提醒、统计看板 |
| `WhatsApp_Templates.xlsx` | WhatsApp/微信/Email 快速回复话术库 |
| `East_MY_Calculator.xlsx` | 东马转运费用计算器（PKG→Kuching/KK + Margin计算） |
| `Quote_Generator.xlsx` | 报价单生成器 — 填黄色格子，自动出报价单 |
| `Document_Agent.xlsx` | 文档Agent — 货物信息→自动生成发票/装箱单/文档清单/船期表 |
| `Followup_Agent.xlsx` | 跟进Agent — 今日跟进/跟进记录/客户分级/消息模板/看板 |
| `Finance_Agent.xlsx` | 财务Agent — 收支记录/应收应付/月度报表/财务看板 |

---

## 📋 报价流程（3分钟出一个报价）

### Step 1: 客户问价
客户通过 WhatsApp / Email / 微信发来询价，你需要确认：
- 运输方式（海运整柜/拼柜/空运）
- 起运地 & 目的地
- 货物信息（品名、重量、体积、柜型）
- 时间要求

### Step 2: 查运价
打开 `Rate_Database.xlsx`，根据航线和柜型查供应商报价：
- 海运整柜 → Sheet "Sea FCL"
- 海运拼柜 → Sheet "Sea LCL"
- 空运 → Sheet "Air Freight"
- 附加费 → Sheet "Surcharges"

### Step 3: 计算报价
在运价基础上加你的 margin（10%-30%），同时加上：
- 附加费（THC、DOC、报关费、拖车费等）
- 保险费（如果客户需要）

### Step 4: 填写报价单
1. 复制 `Quotation_Template.xlsx`，重命名：`Quote_[客户名]_[日期].xlsx`
2. 填写客户信息
3. 填写货物信息
4. 填写费用明细
5. 修改条款和有效期

### Step 5: 发送报价
- **WhatsApp/微信**：截图报价单发过去，或发 Excel 文件
- **Email**：附上 Excel 或 PDF

---

## 💬 快速回复话术

### 收到询价时（WhatsApp/微信）：

```
收到，我帮你查一下价格 👍
请问：
1. 从哪里出？（哪个港口/城市）
2. 到马来西亚哪里？
3. 什么货？大概多重/多大？
4. 整柜还是拼柜？
5. 大概什么时候出？
```

### 发送报价时：

```
Hi [客户名]，

报价如下，有效期7天：

📦 [柜型] [起运地] → [目的地]
💰 USD [总价] (all-in)

包含：海运费、THC、DOC、报关费
不含：拖车费（按实际）、保险（可选）

有问题随时联系我 🙏
```

---

## 🔧 运价更新建议

- 每周跟供应商要一次最新运价
- 更新 `Rate_Database.xlsx` 中的费率和有效期
- 记住：没有固定供应商 = 你的优势是灵活比价
- 拿到新报价先比价，选最优的给客户，自己留足 margin

---

## ⚡ 后续可以加的功能

- [ ] 自动计算 margin 的公式
- [ ] WhatsApp 自动回复 bot
- [ ] 运价趋势跟踪
