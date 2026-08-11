# 酒店下单采购系统

[中文](#中文说明) | [English](#english)

## 中文说明

这是一个面向酒店配送业务的下单、销售单、采购单和利润统计系统。系统从原库存管理项目中独立出来，专门用于酒店客户通过文字清单下单，后台统一生成销售单、采购单，并按客户、商品和品类沉淀价格数据。

## 适用场景

- 酒店在微信群里发送采购清单。
- 后台把清单导入系统，自动识别商品、数量、单位和备注。
- 按客户历史售价生成销售单。
- 按商品分类和供应商习惯导出采购单。
- 打印销售单和采购单。
- 后续统计客户销售额、商品采购量和利润。

## 主要功能

- 清单下单：粘贴微信群文字清单，自动拆分成订单明细。
- 商品标准化：把同类商品统一名称，方便分类、统计和对账。
- 客户售价：按客户、商品、单位保存销售价。
- 销售单导出：支持 Excel 导出，金额列带公式，单价和金额保留两位小数。
- 采购单导出：支持单个客户采购单，也支持多客户合并采购单。
- 水果分单：支持按客户规则把水果单独生成销售单。
- 利润报表：按客户、品类和单品查看销售、进货和利润。
- 网页打印：销售单适配二等分连续纸 `241mm x 140mm`，采购单适配 A4。
- 客户入口：客户可通过独立链接下单、查看记录、复制历史订单。
- 打印助手：Windows 台式机可接收网页后台发送的打印任务。

## 技术栈

- 后端：Python Flask
- 数据库：SQLite
- 前端：原生 HTML、CSS、JavaScript
- Excel：openpyxl
- 部署：Gunicorn、Nginx、systemd

## 本地运行

```bash
cd /Users/donggenyuan/.newmax/workspace/hotel-order-system
python3 -m pip install -r requirements.txt
PORT=5011 python3 server.py
```

浏览器打开：

```text
http://localhost:5011
```

## 生产部署

部署文件放在 `deploy/` 目录：

- `deploy/deploy_vps.sh`：VPS 部署脚本
- `deploy/systemd/hotel-order.service`：systemd 服务模板
- `deploy/nginx/hotel-order.conf`：Nginx 反向代理模板
- `deploy/SECURITY.md`：安全注意事项

默认应用端口是 `127.0.0.1:5011`，建议通过 Nginx 对外访问。

## 不提交到 GitHub 的内容

以下内容已经通过 `.gitignore` 排除：

- `.env`
- `database.db`
- 数据库备份文件
- 客户订单导出文件
- 二维码图片
- 打印助手本机配置
- 临时脚本和临时输出目录

这些文件可能包含客户资料、订单数据、价格数据、服务器地址或访问令牌，不应提交到仓库。

## 迭代规则

每次修改系统后，建议按下面方式提交：

```bash
git status
git add .
git commit -m "修复：清单导入商品名称统一"
git push
```

提交说明要写清楚解决了什么问题，例如：

- `修复：青菜统一为青菜（把）`
- `新增：批量导出销售单按客户分 sheet`
- `优化：合并采购单按标准商品名统计`
- `修复：销售单打印纸张尺寸超出`

这样后面如果出现问题，可以通过 GitHub 查看历史记录，也可以回滚到稳定版本。

## 回滚思路

- 小问题：撤回最近一次提交。
- 大问题：回到某个稳定提交。
- VPS 问题：从 GitHub 拉取稳定版本重新部署。

生产环境回滚前，先备份服务器上的 `database.db`。

## 重要提醒

这个仓库存放的是系统代码，不存放正式业务数据库。正式订单、客户、商品、售价和利润数据以服务器上的 `database.db` 为准，更新和部署前必须先备份数据库。

---

## English

# Hotel Order and Procurement System

This is a hotel-focused order, sales invoice, procurement order, and profit reporting system. It was separated from a general inventory system and adapted for hotel delivery workflows. Hotel customers can submit free-text purchase lists, and the back office can turn them into sales invoices, procurement orders, customer-specific prices, and profit reports.

## Use Cases

- Hotels send purchase lists in WeChat groups.
- Staff paste the text list into the system.
- The system identifies product names, quantities, units, and notes.
- Sales invoices are generated with customer-specific historical prices.
- Procurement orders are exported by product category and supplier workflow.
- Sales invoices and procurement orders can be printed.
- Sales amount, procurement quantity, and profit can be analyzed later.

## Main Features

- Text order import: paste hotel purchase lists and convert them into structured order items.
- Product normalization: standardize similar product names for reporting, classification, and reconciliation.
- Customer pricing: store sales prices by customer, product, and unit.
- Sales invoice export: export Excel files with formulas, two-decimal unit prices, and amount formatting.
- Procurement order export: export single-customer procurement orders or merged multi-customer procurement orders.
- Fruit split rules: generate separate fruit sales invoices according to customer-specific rules.
- Profit reports: review sales, purchase cost, and profit by customer, category, and item.
- Web printing: sales invoice printing supports `241mm x 140mm` half-page continuous paper; procurement orders support A4.
- Customer portal: customers can place orders, view order history, and copy previous orders through dedicated links.
- Print assistant: a Windows desktop tool can receive print jobs from the web back office.

## Tech Stack

- Backend: Python Flask
- Database: SQLite
- Frontend: plain HTML, CSS, and JavaScript
- Excel processing: openpyxl
- Deployment: Gunicorn, Nginx, systemd

## Local Development

```bash
cd /Users/donggenyuan/.newmax/workspace/hotel-order-system
python3 -m pip install -r requirements.txt
PORT=5011 python3 server.py
```

Open in a browser:

```text
http://localhost:5011
```

## Production Deployment

Deployment files are stored in the `deploy/` directory:

- `deploy/deploy_vps.sh`: VPS deployment script
- `deploy/systemd/hotel-order.service`: systemd service template
- `deploy/nginx/hotel-order.conf`: Nginx reverse proxy template
- `deploy/SECURITY.md`: security notes

The application listens on `127.0.0.1:5011` by default. Public access should go through Nginx.

## Files Not Committed to GitHub

The following files are excluded by `.gitignore`:

- `.env`
- `database.db`
- database backup files
- exported customer order files
- QR code images
- local print assistant configuration
- temporary scripts and output folders

These files may contain customer data, order records, price data, server addresses, or access tokens, so they should not be committed to the repository.

## Iteration Workflow

After each change, use clear commit messages:

```bash
git status
git add .
git commit -m "Fix: normalize product names during text import"
git push
```

Good commit message examples:

- `Fix: normalize 青菜 as 青菜（把）`
- `Add: export sales invoices by customer sheet`
- `Improve: merge procurement orders by standardized product names`
- `Fix: sales invoice print size overflow`

Clear commits make it easier to review history, understand why a change was made, and roll back safely.

## Rollback Strategy

- Small issue: revert the latest commit.
- Larger issue: reset to a known stable commit.
- VPS issue: redeploy a stable version from GitHub.

Always back up the production `database.db` before rolling back or redeploying.

## Important Note

This repository stores application code only. It does not store the production business database. Production orders, customers, products, prices, and profit data are stored in the server-side `database.db`. Always back up the database before updating or deploying the system.
