# 酒店订单管理系统

这是从原学校库存系统独立复制出来的酒店专用版本，原系统不受影响。

当前系统里的客户、商品、售价、分类、酒店清单识别逻辑都随这个独立版本一起保留。第一次部署到 VPS 时会带上当前 `database.db`，同事用 IP 打开后看到的是同一套酒店数据。

## 主要功能

- 酒店清单下单：粘贴微信群文字清单，自动生成销售单。
- 销售单导出：支持 Excel 导出，金额列带公式，单价/金额保留 2 位小数。
- 采购订单导出：支持单张销售单导出采购单，也支持多家酒店合并导出采购单。
- 利润报表：按客户、品类、单品查看和导出利润。
- 打印销售单：网页打印尺寸按二等分三联纸 `241mm x 140mm`。

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

## VPS 默认端口

应用服务默认监听 `127.0.0.1:5011`，正式对外访问建议通过 Nginx 反向代理。

部署脚本再次运行时，会先备份服务器上的数据库到 `/opt/hotel-order-backups`，并保留服务器正在使用的 `database.db`，避免覆盖同事已经录入的新订单。
