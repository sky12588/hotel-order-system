#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/opt/hotel-order-system"
APP_USER="hotelapp"
BACKUP_DIR="/opt/hotel-order-backups"

if [ "$(id -u)" -ne 0 ]; then
  echo "请用 root 执行：sudo bash deploy/deploy_vps.sh"
  exit 1
fi

apt-get update
apt-get install -y python3 python3-venv python3-pip nginx ufw rsync

if ! id "$APP_USER" >/dev/null 2>&1; then
  useradd --system --create-home --shell /usr/sbin/nologin "$APP_USER"
fi

mkdir -p "$APP_DIR" "$BACKUP_DIR"

if [ -f "$APP_DIR/database.db" ]; then
  ts="$(date +%Y%m%d_%H%M%S)"
  cp "$APP_DIR/database.db" "$BACKUP_DIR/database_$ts.db"
  cp "$APP_DIR/database.db" /tmp/hotel-order-database.db.keep
fi

rsync -a --delete \
  --exclude='.venv' \
  --exclude='deploy' \
  --exclude='*.db-shm' \
  --exclude='*.db-wal' \
  ./ "$APP_DIR"/

if [ -f /tmp/hotel-order-database.db.keep ]; then
  mv /tmp/hotel-order-database.db.keep "$APP_DIR/database.db"
fi

chown -R "$APP_USER:$APP_USER" "$APP_DIR"
chown -R "$APP_USER:$APP_USER" "$BACKUP_DIR"

python3 -m venv "$APP_DIR/.venv"
"$APP_DIR/.venv/bin/pip" install --upgrade pip
"$APP_DIR/.venv/bin/pip" install -r "$APP_DIR/requirements.txt"

cp deploy/systemd/hotel-order.service /etc/systemd/system/hotel-order.service
cp deploy/nginx/hotel-order.conf /etc/nginx/sites-available/hotel-order.conf
ln -sfn /etc/nginx/sites-available/hotel-order.conf /etc/nginx/sites-enabled/hotel-order.conf
rm -f /etc/nginx/sites-enabled/default

nginx -t
systemctl daemon-reload
systemctl enable --now hotel-order
systemctl reload nginx

ufw allow OpenSSH
ufw allow 80/tcp
ufw --force enable

echo "部署完成。访问：http://服务器IP/"
