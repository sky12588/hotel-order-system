#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Migrate legacy inventory data.json into the v2 SQLite database."""

import json
import os
import shutil
import sqlite3
import sys
from datetime import date, datetime


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OLD_DATA_PATH = os.path.join(os.path.dirname(BASE_DIR), 'inventory', 'data.json')
DB_PATH = os.path.join(BASE_DIR, 'database.db')
BACKUP_DIR = os.path.join(BASE_DIR, 'migration_backups')


def now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def as_float(value, default=0):
    try:
        return float(value if value not in (None, '') else default)
    except (TypeError, ValueError):
        return float(default)


def clean_number(value):
    number = as_float(value)
    return 0 if abs(number) < 0.000001 else number


def normalize_date(value):
    if not value:
        return date.today().strftime('%Y-%m-%d')
    text = str(value).strip().replace('/', '-')
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(text, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    parts = text.split('-')
    if len(parts) == 3 and all(part.isdigit() for part in parts):
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}-{int(parts[2]):02d}"
    return text


def json_array(value):
    if isinstance(value, list):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            return json.dumps(parsed if isinstance(parsed, list) else [], ensure_ascii=False)
        except Exception:
            return '[]'
    return '[]'


def first(data, *keys, default=''):
    for key in keys:
        if key in data and data[key] not in (None, ''):
            return data[key]
    return default


def product_snapshot(product_map, product_id):
    product = product_map.get(product_id, {})
    return product.get('name', ''), product.get('spec', ''), product.get('unit', '')


def backup_current_database():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'database_before_old_import_{stamp}.db')
    shutil.copy2(DB_PATH, backup_path)
    return backup_path


def main():
    source_path = sys.argv[1] if len(sys.argv) > 1 else OLD_DATA_PATH
    with open(source_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    backup_path = backup_current_database()
    conn = sqlite3.connect(DB_PATH)
    conn.execute('PRAGMA foreign_keys = ON')

    counts = {
        'products': len(data.get('products', [])),
        'suppliers': len(data.get('suppliers', [])),
        'customers': len(data.get('customers', [])),
        'purchases': len(data.get('purchases', [])),
        'outbounds': len(data.get('outbounds', [])),
        'sales': len(data.get('sales', [])),
        'finances': len(data.get('finances', [])),
        'purchase_items': 0,
        'outbound_items': 0,
        'sales_items': 0
    }

    product_map = {p.get('id'): p for p in data.get('products', [])}

    try:
        tables = [
            'finances', 'sales_items', 'sales', 'outbound_items', 'outbounds',
            'purchase_items', 'purchases', 'customers', 'suppliers', 'products'
        ]
        for table in tables:
            conn.execute(f'DELETE FROM {table}')

        for p in data.get('products', []):
            conn.execute('''
                INSERT INTO products
                    (id, code, name, spec, unit, cost, stock, alert_line, departments, last_sale_price, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                p['id'], p['code'], p['name'], p.get('spec', ''), p.get('unit', ''),
                as_float(p.get('cost')), clean_number(p.get('stock')),
                as_float(first(p, 'alert_line', 'alertLine', default=10), 10),
                json_array(p.get('departments')),
                as_float(first(p, 'last_sale_price', 'lastSalePrice')),
                p.get('created_at') or now_str()
            ))

        for s in data.get('suppliers', []):
            conn.execute('''
                INSERT INTO suppliers (id, code, name, contact, phone, address, note, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                s['id'], s.get('code') or s['id'], s.get('name', ''),
                s.get('contact', ''), s.get('phone', ''), s.get('address', ''),
                s.get('note', ''), s.get('created_at') or now_str()
            ))

        for c in data.get('customers', []):
            conn.execute('''
                INSERT INTO customers (id, code, name, company, phone, address, note, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                c['id'], c.get('code') or c['id'], c.get('name', ''),
                c.get('company', ''), c.get('phone', ''), c.get('address', ''),
                c.get('note', ''), c.get('created_at') or now_str()
            ))

        for p in data.get('purchases', []):
            conn.execute('''
                INSERT INTO purchases (id, no, supplier_id, supplier_name, date, total, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                p['id'], p['no'], first(p, 'supplier_id', 'supplierId'),
                first(p, 'supplier_name', 'supplierName', 'supplier'),
                normalize_date(p.get('date')), as_float(p.get('total')), p.get('created_at') or now_str()
            ))
            for index, item in enumerate(p.get('items', []), start=1):
                product_id = first(item, 'product_id', 'productId')
                name, spec, unit = product_snapshot(product_map, product_id)
                quantity = as_float(item.get('quantity'))
                price = as_float(item.get('price'))
                conn.execute('''
                    INSERT INTO purchase_items
                        (id, purchase_id, product_id, product_name, product_spec, product_unit, quantity, price, subtotal)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    item.get('id') or f"{p['id']}-item-{index}", p['id'], product_id,
                    name, spec, unit, quantity, price, as_float(item.get('subtotal'), quantity * price)
                ))
                counts['purchase_items'] += 1

        for o in data.get('outbounds', []):
            conn.execute('''
                INSERT INTO outbounds (id, no, department, person, date, total, purchase_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                o['id'], o['no'], o.get('department', ''), o.get('person', ''),
                normalize_date(o.get('date')), as_float(o.get('total')),
                first(o, 'purchase_id', 'purchaseId'), o.get('created_at') or now_str()
            ))
            for index, item in enumerate(o.get('items', []), start=1):
                product_id = first(item, 'product_id', 'productId')
                name, spec, unit = product_snapshot(product_map, product_id)
                quantity = as_float(item.get('quantity'))
                price = as_float(item.get('price'))
                conn.execute('''
                    INSERT INTO outbound_items
                        (id, outbound_id, product_id, product_name, product_spec, product_unit, quantity, price, subtotal)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    item.get('id') or f"{o['id']}-item-{index}", o['id'], product_id,
                    name, spec, unit, quantity, price, as_float(item.get('subtotal'), quantity * price)
                ))
                counts['outbound_items'] += 1

        for s in data.get('sales', []):
            conn.execute('''
                INSERT INTO sales
                    (id, no, company, customer, phone, date, total, show_handlers, handler, issuer, purchase_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                s['id'], s['no'], s.get('company', ''), s.get('customer', ''),
                s.get('phone', ''), normalize_date(s.get('date')), as_float(s.get('total')),
                1 if first(s, 'show_handlers', 'showHandlers', default=0) else 0,
                s.get('handler', ''), s.get('issuer', ''), first(s, 'purchase_id', 'purchaseId'),
                s.get('created_at') or now_str()
            ))
            for index, item in enumerate(s.get('items', []), start=1):
                product_id = first(item, 'product_id', 'productId')
                name, spec, unit = product_snapshot(product_map, product_id)
                quantity = as_float(item.get('quantity'))
                price = as_float(item.get('price'))
                conn.execute('''
                    INSERT INTO sales_items
                        (id, sale_id, product_id, product_name, product_spec, product_unit, quantity, price, subtotal)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    item.get('id') or f"{s['id']}-item-{index}", s['id'], product_id,
                    name, spec, unit, quantity, price, as_float(item.get('subtotal'), quantity * price)
                ))
                counts['sales_items'] += 1

        for f in data.get('finances', []):
            conn.execute('''
                INSERT INTO finances (id, date, type, category, amount, note, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                f['id'], normalize_date(f.get('date')), f.get('type', ''),
                f.get('category', ''), as_float(f.get('amount')), f.get('note', ''),
                f.get('created_at') or now_str()
            ))

        foreign_key_errors = conn.execute('PRAGMA foreign_key_check').fetchall()
        if foreign_key_errors:
            raise RuntimeError(f'foreign key check failed: {foreign_key_errors}')

        integrity = conn.execute('PRAGMA integrity_check').fetchone()[0]
        if integrity != 'ok':
            raise RuntimeError(f'integrity check failed: {integrity}')

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    report = {
        'migratedAt': now_str(),
        'oldDataPath': source_path,
        'databasePath': DB_PATH,
        'backupPath': backup_path,
        'counts': counts
    }
    report_path = os.path.join(BACKUP_DIR, 'last_migration_report.json')
    with open(report_path, 'w', encoding='utf-8') as file:
        json.dump(report, file, ensure_ascii=False, indent=2)

    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
