# 从本机上传到 VPS

服务器 IP：`服务器IP`

需要先知道 SSH 用户名。下面用 `root` 举例，如果你的用户名不是 root，把命令里的 `root` 换掉。

## 1. 上传部署包

```bash
scp /Users/donggenyuan/.newmax/workspace/hotel-order-system-部署包.tar.gz root@服务器IP:/tmp/
```

## 2. 登录服务器

```bash
ssh root@服务器IP
```

## 3. 解压并部署

```bash
cd /tmp
tar -xzf hotel-order-system-部署包.tar.gz
cd hotel-order-system
bash deploy/deploy_vps.sh
```

## 4. 改登录密码

部署脚本默认会创建网页登录保护，默认账号是 `admin`，默认密码是 `ChangeThisStrongPassword`。

上线后必须修改：

```bash
nano /etc/systemd/system/hotel-order.service
```

把这两行改成自己的账号密码：

```text
Environment=HOTEL_AUTH_USER=admin
Environment=HOTEL_AUTH_PASSWORD=ChangeThisStrongPassword
```

然后重启：

```bash
systemctl daemon-reload
systemctl restart hotel-order
```

## 5. 访问

```text
http://服务器IP/
```

浏览器会弹出登录框，输入上面设置的账号密码。
