# VPS 安全建议

## 必做

- 使用普通服务用户 `hotelapp` 运行程序，不用 root 跑应用。
- 应用只监听 `127.0.0.1:5011`，外网只开放 Nginx 的 80/443。
- 防火墙只开放 SSH、HTTP、HTTPS。
- 数据库文件定期备份：`/opt/hotel-order-system/database.db`。

## SSH

- 建议开启 SSH 密钥登录。
- 如果暂时使用密码，密码要足够复杂，上线后尽快改成密钥。
- 不建议长期允许 root 密码登录。

## HTTPS

如果有域名，建议安装证书：

```bash
apt-get install -y certbot python3-certbot-nginx
certbot --nginx -d 你的域名
```

如果只用 IP 测试，可以先用 HTTP。

