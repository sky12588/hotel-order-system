/* ===== 酒店订单管理系统 ===== */
/* 前端JavaScript - 模块化架构 */

const API_BASE = window.location.protocol === 'file:' ? 'http://localhost:5011' : window.location.origin;

// ==================== API工具 ====================
const API = {
    async get(endpoint) {
        const res = await fetch(`${API_BASE}/api/${endpoint}`);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
    },
    async post(endpoint, data) {
        const res = await fetch(`${API_BASE}/api/${endpoint}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || `HTTP ${res.status}`);
        return result;
    },
    async delete(endpoint) {
        const res = await fetch(`${API_BASE}/api/${endpoint}`, { method: 'DELETE' });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || `HTTP ${res.status}`);
        return result;
    },
    async upload(endpoint, formData) {
        const res = await fetch(`${API_BASE}/api/${endpoint}`, {
            method: 'POST',
            body: formData
        });
        const text = await res.text();
        let result;
        try {
            result = JSON.parse(text);
        } catch (e) {
            throw new Error(`接口返回异常（HTTP ${res.status}），请重启进销存服务后再试`);
        }
        if (!res.ok) throw new Error(result.error || `HTTP ${res.status}`);
        return result;
    }
};

// ==================== 全局状态 ====================
const Store = {
    products: [],
    purchases: [],
    outbounds: [],
    sales: [],
    suppliers: [],
    customers: [],
    finances: [],
    _endpoints: {
        products: 'products',
        suppliers: 'suppliers',
        customers: 'customers',
        purchases: 'purchases',
        outbounds: 'outbounds',
        sales: 'sales?summary=1',
        finances: 'finances'
    },
    _maps: {},
    _loaded: new Set(),

    async loadAll() {
        await this.refresh(Object.keys(this._endpoints));
    },

    async ensure(keys) {
        const names = Array.isArray(keys) ? keys : [keys];
        const missing = names.filter(name => !this._loaded.has(name));
        if (missing.length > 0) {
            await this.refresh(missing);
        }
    },

    async refresh(keys) {
        const names = Array.isArray(keys) ? keys : [keys];
        const results = await Promise.all(names.map(name => API.get(this._endpoints[name])));
        names.forEach((name, index) => {
            this[name] = results[index];
            this._loaded.add(name);
        });
        this.rebuildIndexes(names);
    },

    rebuildIndexes(keys = Object.keys(this._endpoints)) {
        keys.forEach(name => {
            this._maps[name] = new Map((this[name] || []).map(item => [item.id, item]));
        });
    },

    find(name, id) {
        return this._maps[name]?.get(id);
    },

    upsert(name, item) {
        const list = this[name] || [];
        const index = list.findIndex(row => row.id === item.id);
        if (index >= 0) {
            list[index] = { ...list[index], ...item };
        } else {
            list.push(item);
        }
        this.rebuildIndexes([name]);
    }
};

// ==================== 工具函数 ====================
const Utils = {
    formatMoney(amount) {
        return '¥' + Number(amount || 0).toFixed(2);
    },
    formatQty(value) {
        const number = Number(value || 0);
        return Number.isInteger(number) ? String(number) : number.toFixed(2);
    },
    formatDate(dateStr) {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    },
    today() {
        return this.formatDate(new Date());
    },
    tomorrow() {
        const d = new Date();
        d.setDate(d.getDate() + 1);
        return this.formatDate(d);
    },
    generateNextCode(prefix, items, key = 'code') {
        const nums = items.map(i => {
            const m = i[key].match(new RegExp('^' + prefix + '(\\d+)$'));
            return m ? parseInt(m[1]) : 0;
        }).filter(n => n > 0);
        const max = nums.length > 0 ? Math.max(...nums) : 0;
        return prefix + String(max + 1).padStart(3, '0');
    },
    downloadFile(content, filename, type = 'text/csv') {
        const blob = new Blob([content], { type });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
    },
    debounce(fn, delay = 250) {
        let timer = null;
        return (...args) => {
            clearTimeout(timer);
            timer = setTimeout(() => fn(...args), delay);
        };
    },
    toast(message, type = 'success') {
        const toast = document.getElementById('toast');
        toast.textContent = message;
        toast.className = 'toast show ' + type;
        setTimeout(() => toast.classList.remove('show'), 3000);
    },
    numberToChinese(num) {
        const cnNums = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
        const cnUnits = ['', '拾', '佰', '仟'];
        const cnBigUnits = ['', '万', '亿'];
        const numStr = Math.floor(num).toString();
        let result = '';
        let zero = false;
        let unitPos = 0;
        for (let i = numStr.length - 1; i >= 0; i--) {
            const n = parseInt(numStr[i]);
            if (n === 0) {
                if (!zero && result.length > 0) {
                    result = cnNums[0] + result;
                    zero = true;
                }
            } else {
                result = cnNums[n] + cnUnits[unitPos % 4] + result;
                zero = false;
            }
            unitPos++;
            if (unitPos % 4 === 0 && i > 0) {
                result = cnBigUnits[unitPos / 4] + result;
            }
        }
        return result.replace(/零+/g, '零').replace(/零$/, '').replace(/^$/, '零') + '元整';
    }
};

// ==================== 模态框管理 ====================
const Modals = {
    open(id) {
        document.getElementById(id).classList.add('show');
        document.body.style.overflow = 'hidden';
    },
    close(id) {
        document.getElementById(id).classList.remove('show');
        document.body.style.overflow = '';
    },
    init() {
        document.querySelectorAll('.modal').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    const allowClose = modal.dataset.clickOutside !== 'false';
                    if (allowClose) {
                        this.close(modal.id);
                    }
                }
            });
        });
    }
};

// ==================== 应用导航 ====================
const App = {
    currentPage: 'sales',
    pageData: {
        dashboard: [],
        products: ['products'],
        suppliers: ['suppliers'],
        purchase: ['products', 'suppliers', 'purchases'],
        outbound: ['products', 'outbounds'],
        sales: ['products', 'customers', 'sales'],
        customers: ['customers'],
        inventory: ['products'],
        'dept-stats': ['products', 'outbounds'],
        'customer-profit': ['customers', 'sales'],
        finance: ['finances'],
        backup: []
    },
    async navigate(page) {
        this.currentPage = page;
        document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
        document.querySelector(`.nav-item[data-page="${page}"]`)?.classList.add('active');
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.getElementById(page).classList.add('active');

        try {
            await Store.ensure(this.pageData[page] || []);
        } catch (e) {
            Utils.toast('加载数据失败：' + e.message, 'error');
            return;
        }

        // 渲染对应页面
        switch (page) {
            case 'dashboard': Dashboard.render(); break;
            case 'products': Products.render(); break;
            case 'suppliers': Suppliers.render(); break;
            case 'purchase': Purchases.render(); break;
            case 'outbound': Outbounds.render(); break;
            case 'sales': Sales.render(); break;
            case 'customers': Customers.render(); break;
            case 'inventory': Inventory.render(); break;
            case 'dept-stats': DeptStats.render(); break;
            case 'customer-profit': CustomerProfit.render(); break;
            case 'finance': Finance.render(); break;
            case 'backup': Backup.render(); break;
        }

        if (window.innerWidth <= 768) {
            this.toggleMobileMenu();
        }
    },
    toggleMobileMenu() {
        document.getElementById('sidebar').classList.toggle('open');
        document.getElementById('mobileOverlay').classList.toggle('show');
    },
    async init() {
        Modals.init();
        await this.navigate('sales');
        Backup.checkAutoBackup();
    }
};

// ==================== 首页概览 ====================
const Dashboard = {
    async render() {
        try {
            const data = await API.get('dashboard');
            document.getElementById('dashMonthSales').textContent = Utils.formatMoney(data.monthSales || 0);
            document.getElementById('dashTodayOrders').textContent = data.todayOrders || 0;
            document.getElementById('dashMonthOrders').textContent = data.monthOrders || 0;
            document.getElementById('dashActiveCustomers').textContent = data.activeCustomers || 0;

            const customerTbody = document.getElementById('dashCustomerStats');
            if (!data.customerStats || data.customerStats.length === 0) {
                customerTbody.innerHTML = '<tr><td colspan="4" class="empty-state">暂无酒店销售记录</td></tr>';
            } else {
                customerTbody.innerHTML = data.customerStats.map(row => `
                    <tr>
                        <td>${row.customer}</td>
                        <td>${Utils.formatMoney(row.monthSales || 0)}</td>
                        <td>${row.monthOrders || 0}</td>
                        <td>${row.lastOrderDate || '-'}</td>
                    </tr>
                `).join('');
            }

            const itemTbody = document.getElementById('dashTopItems');
            if (!data.topItems || data.topItems.length === 0) {
                itemTbody.innerHTML = '<tr><td colspan="4" class="empty-state">暂无物品采购记录</td></tr>';
            } else {
                itemTbody.innerHTML = data.topItems.map(row => `
                    <tr>
                        <td>${row.name}</td>
                        <td>${row.type || '-'}</td>
                        <td>${Utils.formatQty(row.quantity || 0)}${row.unit || ''}</td>
                        <td>${Utils.formatMoney(row.revenue || 0)}</td>
                    </tr>
                `).join('');
            }
        } catch (e) {
            console.error('Dashboard render error:', e);
        }
    }
};

// ==================== 物品管理 ====================
const Products = {
    selectedIds: new Set(),
    categoryOrder: ['蔬菜类', '豆制品类', '水果类', '肉蛋类', '主食面点类', '粮油调料类', '饮品乳品类', '冻货', '饼干', '其他类'],

    render() {
        const tbody = document.getElementById('productList');
        this.renderCategoryControls();
        const filteredProducts = this.getFilteredProducts();
        if (Store.products.length === 0) {
            tbody.innerHTML = '<tr><td colspan="10" class="empty-state">暂无物品，点击右上角添加</td></tr>';
            return;
        }
        if (filteredProducts.length === 0) {
            tbody.innerHTML = '<tr><td colspan="10" class="empty-state">没有符合条件的物品</td></tr>';
            return;
        }
        tbody.innerHTML = filteredProducts.map(p => {
            const category = p.category || '其他类';
            return `
                <tr>
                    <td><input type="checkbox" ${this.selectedIds.has(p.id) ? 'checked' : ''} onchange="Products.toggleSelect('${p.id}')"></td>
                    <td><span class="category-pill">${category}</span></td>
                    <td>${p.code}</td>
                    <td><strong>${p.name}</strong></td>
                    <td>${p.spec || '-'}</td>
                    <td>${p.unit}</td>
                    <td>${p.cost > 0 ? Utils.formatMoney(p.cost) : '<span style="color:#999">未设置</span>'}</td>
                    <td>${p.last_sale_price > 0 ? Utils.formatMoney(p.last_sale_price) : '<span style="color:#999">未设置</span>'}</td>
                    <td>${p.stock}</td>
                    <td>
                        <button class="btn btn-sm" onclick="ProductModal.edit('${p.id}')">编辑</button>
                        <button class="btn btn-sm btn-danger" onclick="Products.delete('${p.id}')">删除</button>
                    </td>
                </tr>
            `;
        }).join('');
    },

    getCategoryRank(category) {
        const index = this.categoryOrder.indexOf(category || '其他类');
        return index >= 0 ? index : 999;
    },

    getProductCategory(p) {
        return p.category || '其他类';
    },

    getFilters() {
        const categoryEl = document.getElementById('productCategoryFilter');
        const searchEl = document.getElementById('productSearchInput');
        return {
            category: categoryEl ? categoryEl.value : '',
            search: searchEl ? searchEl.value.trim().toLowerCase() : ''
        };
    },

    getFilteredProducts() {
        const filters = this.getFilters();
        return [...Store.products]
            .filter(p => {
                const category = this.getProductCategory(p);
                if (filters.category && category !== filters.category) return false;
                if (!filters.search) return true;
                const haystack = [
                    p.code, p.name, p.spec, p.unit, category, p.purchase_type
                ].map(v => String(v || '').toLowerCase()).join(' ');
                return haystack.includes(filters.search);
            })
            .sort((a, b) => {
                const categoryDiff = this.getCategoryRank(this.getProductCategory(a)) - this.getCategoryRank(this.getProductCategory(b));
                if (categoryDiff !== 0) return categoryDiff;
                return String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hans-CN');
            });
    },

    renderCategoryControls() {
        const select = document.getElementById('productCategoryFilter');
        const summary = document.getElementById('productCategorySummary');
        if (!select || !summary) return;

        const currentValue = select.value;
        const counts = Store.products.reduce((acc, product) => {
            const category = this.getProductCategory(product);
            acc[category] = (acc[category] || 0) + 1;
            return acc;
        }, {});
        const categories = Object.keys(counts).sort((a, b) => {
            const diff = this.getCategoryRank(a) - this.getCategoryRank(b);
            return diff !== 0 ? diff : a.localeCompare(b, 'zh-Hans-CN');
        });

        select.innerHTML = '<option value="">全部分类</option>' + categories.map(category => (
            `<option value="${category}" ${category === currentValue ? 'selected' : ''}>${category}（${counts[category]}）</option>`
        )).join('');

        const activeCategory = currentValue || '';
        const totalClass = activeCategory ? '' : ' active';
        summary.innerHTML = [
            `<button type="button" class="category-stat${totalClass}" onclick="Products.setCategoryFilter('')">全部<span>${Store.products.length}</span></button>`,
            ...categories.map(category => {
                const active = category === activeCategory ? ' active' : '';
                return `<button type="button" class="category-stat${active}" onclick="Products.setCategoryFilter('${category}')">${category}<span>${counts[category]}</span></button>`;
            })
        ].join('');
    },

    setCategoryFilter(category) {
        const select = document.getElementById('productCategoryFilter');
        if (select) select.value = category || '';
        this.render();
    },

    onSearchInput: Utils.debounce(() => Products.render(), 180),

    clearFilters() {
        const select = document.getElementById('productCategoryFilter');
        const search = document.getElementById('productSearchInput');
        if (select) select.value = '';
        if (search) search.value = '';
        this.render();
    },

    toggleSelect(id) {
        if (this.selectedIds.has(id)) {
            this.selectedIds.delete(id);
        } else {
            this.selectedIds.add(id);
        }
    },

    toggleSelectAll() {
        const checked = document.getElementById('selectAllProducts').checked;
        const filteredProducts = this.getFilteredProducts();
        if (checked) {
            filteredProducts.forEach(p => this.selectedIds.add(p.id));
        } else {
            filteredProducts.forEach(p => this.selectedIds.delete(p.id));
        }
        this.render();
    },

    async delete(id) {
        if (!confirm('确定删除该物品？')) return;
        try {
            await API.delete(`products/${id}`);
            await Store.refresh('products');
            this.render();
            Dashboard.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    },

    exportSelected() {
        if (this.selectedIds.size === 0) {
            Utils.toast('请先选择要导出的物品', 'error');
            return;
        }
        const selected = Store.products.filter(p => this.selectedIds.has(p.id));
        let csv = '﻿分类,编号,名称,规格,单位,采购价,销售价,库存\n';
        selected.forEach(p => {
            csv += `${p.category || '其他类'},${p.code},${p.name},${p.spec || ''},${p.unit},${p.cost || 0},${p.last_sale_price || 0},${p.stock || 0}\n`;
        });
        Utils.downloadFile(csv, `选中物品_${Utils.today()}.csv`);
        Utils.toast('导出成功');
    }
};

// ==================== 物品模态框 ====================
const ProductModal = {
    open() {
        document.getElementById('productModalTitle').textContent = '添加物品';
        document.getElementById('productForm').reset();
        document.getElementById('productId').value = '';
        document.getElementById('productCode').value = Utils.generateNextCode('A', Store.products);
        document.getElementById('productAlert').value = '10';
        Modals.open('productModal');
    },

    edit(id) {
        const p = Store.find('products', id);
        if (!p) return;
        document.getElementById('productModalTitle').textContent = '编辑物品';
        document.getElementById('productId').value = p.id;
        document.getElementById('productCode').value = p.code;
        document.getElementById('productName').value = p.name;
        document.getElementById('productSpec').value = p.spec || '';
        document.getElementById('productUnit').value = p.unit;
        document.getElementById('productCost').value = p.cost || 0;
        document.getElementById('productStock').value = p.stock || 0;
        document.getElementById('productAlert').value = p.alert_line || 10;
        document.getElementById('productLastSalePrice').value = p.last_sale_price || 0;

        // 设置部门复选框
        document.querySelectorAll('input[name="productDept"]').forEach(cb => {
            cb.checked = (p.departments || []).includes(cb.value);
        });

        Modals.open('productModal');
    },

    async save(e) {
        e.preventDefault();
        const departments = Array.from(document.querySelectorAll('input[name="productDept"]:checked')).map(cb => cb.value);
        const data = {
            id: document.getElementById('productId').value || undefined,
            code: document.getElementById('productCode').value.trim(),
            name: document.getElementById('productName').value.trim(),
            spec: document.getElementById('productSpec').value.trim(),
            unit: document.getElementById('productUnit').value.trim(),
            cost: parseFloat(document.getElementById('productCost').value) || 0,
            stock: parseFloat(document.getElementById('productStock').value) || 0,
            alertLine: parseFloat(document.getElementById('productAlert').value) || 10,
            departments,
            lastSalePrice: parseFloat(document.getElementById('productLastSalePrice').value) || 0
        };

        try {
            await API.post('products', data);
            await Store.refresh('products');
            Modals.close('productModal');
            Products.render();
            Utils.toast(data.id ? '更新成功' : '添加成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 供应商管理 ====================
const Suppliers = {
    render() {
        const tbody = document.getElementById('supplierList');
        if (Store.suppliers.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="empty-state">暂无供应商</td></tr>';
            return;
        }
        tbody.innerHTML = Store.suppliers.map(s => `
            <tr>
                <td>${s.code}</td>
                <td><strong>${s.name}</strong></td>
                <td>${s.contact || '-'}</td>
                <td>${s.phone || '-'}</td>
                <td>${s.address || '-'}</td>
                <td>${s.note || '-'}</td>
                <td>
                    <button class="btn btn-sm" onclick="SupplierModal.edit('${s.id}')">编辑</button>
                    <button class="btn btn-sm btn-danger" onclick="Suppliers.delete('${s.id}')">删除</button>
                </td>
            </tr>
        `).join('');
    },

    async delete(id) {
        if (!confirm('确定删除该供应商？')) return;
        try {
            await API.delete(`suppliers/${id}`);
            await Store.refresh('suppliers');
            this.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

const SupplierModal = {
    open() {
        document.getElementById('supplierModalTitle').textContent = '添加供应商';
        document.getElementById('supplierForm').reset();
        document.getElementById('supplierId').value = '';
        document.getElementById('supplierCode').value = Utils.generateNextCode('S', Store.suppliers);
        Modals.open('supplierModal');
    },

    edit(id) {
        const s = Store.find('suppliers', id);
        if (!s) return;
        document.getElementById('supplierModalTitle').textContent = '编辑供应商';
        document.getElementById('supplierId').value = s.id;
        document.getElementById('supplierCode').value = s.code;
        document.getElementById('supplierName').value = s.name;
        document.getElementById('supplierContact').value = s.contact || '';
        document.getElementById('supplierPhone').value = s.phone || '';
        document.getElementById('supplierAddress').value = s.address || '';
        document.getElementById('supplierNote').value = s.note || '';
        Modals.open('supplierModal');
    },

    async save(e) {
        e.preventDefault();
        const data = {
            id: document.getElementById('supplierId').value || undefined,
            code: document.getElementById('supplierCode').value.trim(),
            name: document.getElementById('supplierName').value.trim(),
            contact: document.getElementById('supplierContact').value.trim(),
            phone: document.getElementById('supplierPhone').value.trim(),
            address: document.getElementById('supplierAddress').value.trim(),
            note: document.getElementById('supplierNote').value.trim()
        };
        try {
            await API.post('suppliers', data);
            await Store.refresh('suppliers');
            Modals.close('supplierModal');
            Suppliers.render();
            Utils.toast(data.id ? '更新成功' : '添加成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 客户管理 ====================
const Customers = {
    render() {
        const tbody = document.getElementById('customerList');
        if (Store.customers.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="empty-state">暂无客户</td></tr>';
            return;
        }
        tbody.innerHTML = Store.customers.map(c => `
            <tr>
                <td>${c.code}</td>
                <td><strong>${c.name}</strong></td>
                <td>${c.company || '-'}</td>
                <td>${c.phone || '-'}</td>
                <td>${c.address || '-'}</td>
                <td>${c.note || '-'}</td>
                <td>
                    <button class="btn btn-sm" onclick="CustomerModal.edit('${c.id}')">编辑</button>
                    <button class="btn btn-sm btn-danger" onclick="Customers.delete('${c.id}')">删除</button>
                </td>
            </tr>
        `).join('');
    },

    async delete(id) {
        if (!confirm('确定删除该客户？')) return;
        try {
            await API.delete(`customers/${id}`);
            await Store.refresh('customers');
            this.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

const CustomerModal = {
    open() {
        document.getElementById('customerModalTitle').textContent = '添加客户';
        document.getElementById('customerForm').reset();
        document.getElementById('customerId').value = '';
        document.getElementById('customerCode').value = Utils.generateNextCode('C', Store.customers);
        Modals.open('customerModal');
    },

    edit(id) {
        const c = Store.find('customers', id);
        if (!c) return;
        document.getElementById('customerModalTitle').textContent = '编辑客户';
        document.getElementById('customerId').value = c.id;
        document.getElementById('customerCode').value = c.code;
        document.getElementById('customerName').value = c.name;
        document.getElementById('customerCompany').value = c.company || '';
        document.getElementById('customerPhone').value = c.phone || '';
        document.getElementById('customerAddress').value = c.address || '';
        document.getElementById('customerNote').value = c.note || '';
        Modals.open('customerModal');
    },

    async save(e) {
        e.preventDefault();
        const data = {
            id: document.getElementById('customerId').value || undefined,
            code: document.getElementById('customerCode').value.trim(),
            name: document.getElementById('customerName').value.trim(),
            company: document.getElementById('customerCompany').value.trim(),
            phone: document.getElementById('customerPhone').value.trim(),
            address: document.getElementById('customerAddress').value.trim(),
            note: document.getElementById('customerNote').value.trim()
        };
        try {
            await API.post('customers', data);
            await Store.refresh('customers');
            Modals.close('customerModal');
            Customers.render();
            Utils.toast(data.id ? '更新成功' : '添加成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 采购入库 ====================
const Purchases = {
    selectedIds: new Set(),

    render() {
        const tbody = document.getElementById('purchaseList');
        if (Store.purchases.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="empty-state">暂无采购记录</td></tr>';
            return;
        }
        tbody.innerHTML = Store.purchases.map(p => `
            <tr>
                <td><input type="checkbox" ${this.selectedIds.has(p.id) ? 'checked' : ''} onchange="Purchases.toggleSelect('${p.id}')"></td>
                <td>${p.no}</td>
                <td>${p.date}</td>
                <td>${p.supplier_name || '-'}</td>
                <td>${p.items ? p.items.length : 0}</td>
                <td>${Utils.formatMoney(p.total)}</td>
                <td>
                    <button class="btn btn-sm" onclick="Printer.printPurchase('${p.id}')">查看</button>
                    <button class="btn btn-sm" onclick="PurchaseModal.edit('${p.id}')">编辑</button>
                    <button class="btn btn-sm" onclick="Printer.printPurchase('${p.id}')">打印</button>
                    <button class="btn btn-sm btn-danger" onclick="Purchases.delete('${p.id}')">删除</button>
                </td>
            </tr>
        `).join('');
    },

    toggleSelect(id) {
        if (this.selectedIds.has(id)) {
            this.selectedIds.delete(id);
        } else {
            this.selectedIds.add(id);
        }
    },

    toggleSelectAll() {
        const checked = document.getElementById('selectAllPurchases').checked;
        if (checked) {
            Store.purchases.forEach(p => this.selectedIds.add(p.id));
        } else {
            this.selectedIds.clear();
        }
        this.render();
    },

    async delete(id) {
        if (!confirm('确定删除该采购单？库存将回退。')) return;
        try {
            await API.delete(`purchases/${id}`);
            await Store.refresh(['products', 'purchases', 'outbounds', 'sales']);
            this.render();
            Dashboard.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    },

    async batchDelete() {
        if (this.selectedIds.size === 0) {
            Utils.toast('请先选择要删除的采购单', 'error');
            return;
        }
        if (!confirm(`确定删除选中的 ${this.selectedIds.size} 个采购单？`)) return;
        try {
            for (const id of this.selectedIds) {
                await API.delete(`purchases/${id}`);
            }
            this.selectedIds.clear();
            await Store.refresh(['products', 'purchases', 'outbounds', 'sales']);
            this.render();
            Dashboard.render();
            Utils.toast('批量删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 采购单模态框 ====================
const PurchaseModal = {
    items: [],

    open() {
        this.reset();
        document.getElementById('purchaseModalTitle').textContent = '新建采购单';
        document.getElementById('purchaseId').value = '';
        document.getElementById('purchaseDate').value = Utils.today();
        this.renderSupplierOptions();
        this.renderCustomerOptions();
        Modals.open('purchaseModal');
    },

    edit(id) {
        const p = Store.find('purchases', id);
        if (!p) return;
        this.reset();
        document.getElementById('purchaseModalTitle').textContent = '编辑采购单';
        document.getElementById('purchaseId').value = p.id;
        document.getElementById('purchaseDate').value = p.date;
        this.renderSupplierOptions(p.supplier_id);
        this.renderCustomerOptions();

        this.items = p.items.map(i => ({
            productId: i.product_id,
            quantity: i.quantity,
            price: i.price,
            productName: i.product_name,
            productSpec: i.product_spec,
            productUnit: i.product_unit
        }));
        this.renderItems();
        Modals.open('purchaseModal');
    },

    reset() {
        this.items = [];
        document.getElementById('purchaseForm').reset();
        document.getElementById('autoOutboundCheck').checked = false;
        document.getElementById('autoOutboundOptions').style.display = 'none';
        document.getElementById('autoSalesCheck').checked = false;
        document.getElementById('autoSalesOptions').style.display = 'none';
        this.renderItems();
    },

    renderSupplierOptions(selectedId = '') {
        const select = document.getElementById('purchaseSupplier');
        select.innerHTML = '<option value="">选择供应商</option>' +
            Store.suppliers.map(s => `<option value="${s.id}" ${s.id === selectedId ? 'selected' : ''}>${s.name}</option>`).join('');
    },

    renderCustomerOptions() {
        const select = document.getElementById('autoSalesCustomer');
        select.innerHTML = '<option value="">选择客户</option>' +
            Store.customers.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
    },

    renderItems() {
        const tbody = document.getElementById('purchaseItemsBody');
        if (this.items.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px;">点击"添加物品"按钮添加采购物品</td></tr>';
        } else {
            tbody.innerHTML = this.items.map((item, idx) => `
                <tr>
                    <td><strong>${item.productName || '未知'}</strong></td>
                    <td>${item.productSpec || '-'}</td>
                    <td>${item.productUnit || '-'}</td>
                    <td><input type="number" step="0.01" min="0.01" value="${item.quantity}" onchange="PurchaseModal.updateItem(${idx}, 'quantity', this.value)" style="width:70px;"></td>
                    <td><input type="number" step="0.01" min="0" value="${item.price}" onchange="PurchaseModal.updateItem(${idx}, 'price', this.value)" style="width:80px;"></td>
                    <td>${Utils.formatMoney(item.quantity * item.price)}</td>
                    <td><button type="button" class="btn btn-sm btn-danger" onclick="PurchaseModal.removeItem(${idx})">删除</button></td>
                </tr>
            `).join('');
        }
        const total = this.items.reduce((sum, i) => sum + i.quantity * i.price, 0);
        document.getElementById('purchaseTotal').textContent = Utils.formatMoney(total);

        // 刷新自动出库价格设置
        if (document.getElementById('autoOutboundCheck').checked) {
            this.renderAutoOutboundPrices();
        }
        if (document.getElementById('autoSalesCheck').checked) {
            this.renderAutoSalesPrices();
        }
    },

    updateItem(idx, field, value) {
        this.items[idx][field] = parseFloat(value) || 0;
        this.renderItems();
    },

    removeItem(idx) {
        this.items.splice(idx, 1);
        this.renderItems();
    },

    showProductSelector() {
        const container = document.getElementById('purchaseProductSelectorList');
        document.getElementById('purchaseProductSearch').value = '';
        this._availableProducts = Store.products.filter(p => !this.items.some(i => i.productId === p.id));
        this.renderProductCards(this._availableProducts);
        Modals.open('purchaseProductSelector');
    },

    renderProductCards(productList) {
        const container = document.getElementById('purchaseProductSelectorList');
        if (productList.length === 0) {
            container.innerHTML = '<div style="text-align:center;color:#999;padding:40px;">没有可添加的物品</div>';
            return;
        }
        container.innerHTML = productList.map(p => `
            <div class="product-select-card" data-id="${p.id}" onclick="this.classList.toggle('selected')">
                <div style="font-weight:500;color:#333;">${p.name}</div>
                <div style="font-size:12px;color:#888;margin-top:4px;">${p.spec || '-'} | ${p.unit} | 当前单价 ¥${(p.cost || 0).toFixed(2)}</div>
                <div style="margin-top:8px;">
                    <label style="font-size:12px;color:#666;">数量：</label>
                    <input type="number" class="qty-input" step="0.01" min="0.01" value="1" style="width:70px;padding:4px;border:1px solid #ddd;border-radius:4px;" onclick="event.stopPropagation()">
                </div>
            </div>
        `).join('');
    },

    filterProducts() {
        const search = document.getElementById('purchaseProductSearch').value.toLowerCase().trim();
        if (!search) {
            this.renderProductCards(this._availableProducts);
            return;
        }
        const filtered = this._availableProducts.filter(p =>
            p.name.toLowerCase().includes(search) ||
            p.code.toLowerCase().includes(search) ||
            (p.spec && p.spec.toLowerCase().includes(search))
        );
        this.renderProductCards(filtered);
    },

    confirmProductSelect() {
        const selected = document.querySelectorAll('#purchaseProductSelectorList .product-select-card.selected');
        selected.forEach(card => {
            const productId = card.dataset.id;
            const qty = parseFloat(card.querySelector('.qty-input').value) || 1;
            const product = Store.find('products', productId);
            if (product) {
                this.items.push({
                    productId: productId,
                    quantity: qty,
                    price: product.cost || 0,
                    productName: product.name,
                    productSpec: product.spec,
                    productUnit: product.unit
                });
            }
        });
        Modals.close('purchaseProductSelector');
        this.renderItems();
    },

    toggleAutoOutbound() {
        const checked = document.getElementById('autoOutboundCheck').checked;
        document.getElementById('autoOutboundOptions').style.display = checked ? 'block' : 'none';
        if (checked) this.renderAutoOutboundPrices();
    },

    renderAutoOutboundPrices() {
        const container = document.getElementById('autoOutboundPrices');
        if (this.items.length === 0) {
            container.innerHTML = '<p style="color:#999;font-size:12px;">请先添加采购物品</p>';
            return;
        }
        container.innerHTML = this.items.map((item, idx) => `
            <div style="display:flex;align-items:center;gap:10px;padding:6px 10px;background:#fff;border:1px solid #e9ecef;border-radius:4px;margin-bottom:6px;">
                <span style="flex:1;font-size:13px;color:#333;">${item.productName || '未知'}</span>
                <span style="font-size:12px;color:#888;white-space:nowrap;">采购价 ¥${item.price.toFixed(2)}</span>
                <span style="font-size:12px;color:#888;">→</span>
                <input type="number" id="outboundPrice_${idx}" step="0.01" min="0" value="${item.price}"
                    style="width:80px;padding:4px 8px;border:1px solid #d9d9d9;border-radius:4px;font-size:13px;"
                    placeholder="出库单价">
                <span style="font-size:12px;color:#28a745;white-space:nowrap;">利润 ¥<span id="outboundProfit_${idx}">0.00</span></span>
            </div>
        `).join('');

        // 绑定利润计算
        this.items.forEach((item, idx) => {
            const input = document.getElementById(`outboundPrice_${idx}`);
            if (input) {
                const calc = () => {
                    const price = parseFloat(input.value) || 0;
                    const profit = (price - item.price) * item.quantity;
                    const el = document.getElementById(`outboundProfit_${idx}`);
                    if (el) {
                        el.textContent = profit.toFixed(2);
                        el.style.color = profit >= 0 ? '#28a745' : '#e74c3c';
                    }
                };
                input.addEventListener('input', calc);
                calc();
            }
        });
    },

    toggleAutoSales() {
        const checked = document.getElementById('autoSalesCheck').checked;
        document.getElementById('autoSalesOptions').style.display = checked ? 'block' : 'none';
        if (checked) this.renderAutoSalesPrices();
    },

    onSalesCustomerChange() {
        const customerId = document.getElementById('autoSalesCustomer').value;
        if (!customerId) return;
        const customer = Store.find('customers', customerId);
        if (customer) {
            document.getElementById('autoSalesCompany').value = customer.company || '';
            document.getElementById('autoSalesCustomerName').value = customer.name;
            document.getElementById('autoSalesPhone').value = customer.phone || '';
        }
    },

    renderAutoSalesPrices() {
        const container = document.getElementById('autoSalesPrices');
        if (this.items.length === 0) {
            container.innerHTML = '<p style="color:#999;font-size:12px;">请先添加采购物品</p>';
            return;
        }
        container.innerHTML = this.items.map((item, idx) => {
            const product = Store.find('products', item.productId);
            const defaultPrice = product && product.last_sale_price ? product.last_sale_price : item.price;
            return `
                <div style="display:flex;align-items:center;gap:10px;padding:6px 10px;background:#fff;border:1px solid #e9ecef;border-radius:4px;margin-bottom:6px;">
                    <span style="flex:1;font-size:13px;color:#333;">${item.productName || '未知'}</span>
                    <span style="font-size:12px;color:#888;white-space:nowrap;">采购价 ¥${item.price.toFixed(2)}</span>
                    <span style="font-size:12px;color:#888;">→</span>
                    <input type="number" id="salesPrice_${idx}" step="0.01" min="0" value="${defaultPrice}"
                        style="width:80px;padding:4px 8px;border:1px solid #d9d9d9;border-radius:4px;font-size:13px;"
                        placeholder="销售单价">
                    <span style="font-size:12px;color:#28a745;white-space:nowrap;">利润 ¥<span id="salesProfit_${idx}">0.00</span></span>
                </div>
            `;
        }).join('');

        this.items.forEach((item, idx) => {
            const input = document.getElementById(`salesPrice_${idx}`);
            if (input) {
                const calc = () => {
                    const price = parseFloat(input.value) || 0;
                    const profit = (price - item.price) * item.quantity;
                    const el = document.getElementById(`salesProfit_${idx}`);
                    if (el) {
                        el.textContent = profit.toFixed(2);
                        el.style.color = profit >= 0 ? '#28a745' : '#e74c3c';
                    }
                };
                input.addEventListener('input', calc);
                calc();
            }
        });
    },

    async save(e) {
        e.preventDefault();
        if (this.items.length === 0) {
            Utils.toast('请至少添加一个物品', 'error');
            return;
        }

        const supplierSelect = document.getElementById('purchaseSupplier');
        const supplierId = supplierSelect.value;
        const supplierName = supplierSelect.options[supplierSelect.selectedIndex]?.text || '';

        // 收集出库单价
        const outboundPrices = {};
        if (document.getElementById('autoOutboundCheck').checked) {
            this.items.forEach((_, idx) => {
                const input = document.getElementById(`outboundPrice_${idx}`);
                if (input) outboundPrices[`outboundPrice_${idx}`] = input.value;
            });
        }

        // 收集销售单价
        const salesPrices = {};
        if (document.getElementById('autoSalesCheck').checked) {
            this.items.forEach((_, idx) => {
                const input = document.getElementById(`salesPrice_${idx}`);
                if (input) salesPrices[`salesPrice_${idx}`] = input.value;
            });
        }

        const data = {
            id: document.getElementById('purchaseId').value || undefined,
            supplierId,
            supplierName,
            date: document.getElementById('purchaseDate').value,
            items: this.items.map(i => ({
                productId: i.productId,
                quantity: i.quantity,
                price: i.price
            })),
            autoOutbound: document.getElementById('autoOutboundCheck').checked,
            autoOutboundDept: document.getElementById('autoOutboundDept').value,
            autoOutboundPerson: document.getElementById('autoOutboundPerson').value,
            outboundPrices,
            autoSales: document.getElementById('autoSalesCheck').checked,
            autoSalesCustomer: document.getElementById('autoSalesCustomerName').value,
            autoSalesCompany: document.getElementById('autoSalesCompany').value,
            autoSalesPhone: document.getElementById('autoSalesPhone').value,
            autoSalesDate: document.getElementById('autoSalesDate').value,
            salesPrices
        };

        try {
            const result = await API.post('purchases', data);
            await Store.refresh(['products', 'purchases', 'outbounds', 'sales']);
            Modals.close('purchaseModal');
            Purchases.render();
            Dashboard.render();
            if (result.outbound) {
                Utils.toast(`✅ 入库成功，已自动生成 ${result.outbound.dept} 的出库单（${result.outbound.no}）`);
            } else if (result.sales) {
                Utils.toast(`✅ 入库成功，已自动生成销售出库单（${result.sales.no}）`);
            } else {
                Utils.toast('保存成功');
            }
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 快速添加物品 ====================
const QuickProductModal = {
    open() {
        document.getElementById('quickProductForm').reset();
        document.getElementById('quickProductCode').value = Utils.generateNextCode('A', Store.products);
        document.getElementById('quickProductAlert').value = '10';
        Modals.open('quickProductModal');
    },

    async save(e) {
        e.preventDefault();
        const data = {
            code: document.getElementById('quickProductCode').value.trim(),
            name: document.getElementById('quickProductName').value.trim(),
            spec: document.getElementById('quickProductSpec').value.trim(),
            unit: document.getElementById('quickProductUnit').value.trim(),
            cost: parseFloat(document.getElementById('quickProductCost').value) || 0,
            stock: 0,
            alertLine: parseFloat(document.getElementById('quickProductAlert').value) || 10,
            departments: [],
            lastSalePrice: 0
        };

        try {
            const result = await API.post('products', data);
            await Store.refresh('products');
            Modals.close('quickProductModal');

            // 自动添加到采购单
            const newProduct = Store.find('products', result.id);
            if (newProduct) {
                PurchaseModal.items.push({
                    productId: newProduct.id,
                    quantity: 1,
                    price: newProduct.cost || 0,
                    productName: newProduct.name,
                    productSpec: newProduct.spec,
                    productUnit: newProduct.unit
                });
                PurchaseModal.renderItems();
            }
            Utils.toast('添加成功并已选中');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 部门领用 ====================
const Outbounds = {
    selectedIds: new Set(),

    render() {
        const deptFilter = document.getElementById('outboundDeptFilter').value;
        let filtered = Store.outbounds;
        if (deptFilter) {
            filtered = filtered.filter(o => o.department === deptFilter);
        }

        const tbody = document.getElementById('outboundList');
        if (filtered.length === 0) {
            this.selectedIds.clear();
            const selectAll = document.getElementById('selectAllOutbounds');
            if (selectAll) selectAll.checked = false;
            tbody.innerHTML = '<tr><td colspan="7" class="empty-state">暂无领用记录</td></tr>';
            return;
        }
        const visibleIds = new Set(filtered.map(o => o.id));
        this.selectedIds = new Set([...this.selectedIds].filter(id => visibleIds.has(id)));
        const selectAll = document.getElementById('selectAllOutbounds');
        if (selectAll) selectAll.checked = filtered.length > 0 && filtered.every(o => this.selectedIds.has(o.id));
        tbody.innerHTML = filtered.map(o => `
            <tr>
                <td><input type="checkbox" ${this.selectedIds.has(o.id) ? 'checked' : ''} onchange="Outbounds.toggleSelect('${o.id}', this.checked)"></td>
                <td>${o.no}</td>
                <td>${o.date}</td>
                <td><span class="status status-normal">${o.department}</span></td>
                <td>${o.items ? o.items.length : 0}</td>
                <td>${Utils.formatMoney(o.total)}</td>
                <td>
                    <button class="btn btn-sm" onclick="OutboundDetailModal.open('${o.id}')">查看</button>
                    <button class="btn btn-sm" onclick="OutboundModal.edit('${o.id}')">编辑</button>
                    <button class="btn btn-sm" onclick="Printer.printOutbound('${o.id}')">打印</button>
                    <button class="btn btn-sm btn-danger" onclick="Outbounds.delete('${o.id}')">删除</button>
                </td>
            </tr>
        `).join('');
    },

    toggleSelect(id, checked) {
        if (checked) {
            this.selectedIds.add(id);
        } else {
            this.selectedIds.delete(id);
        }
        const deptFilter = document.getElementById('outboundDeptFilter').value;
        const visible = deptFilter ? Store.outbounds.filter(o => o.department === deptFilter) : Store.outbounds;
        const selectAll = document.getElementById('selectAllOutbounds');
        if (selectAll) selectAll.checked = visible.length > 0 && visible.every(o => this.selectedIds.has(o.id));
    },

    toggleSelectAll() {
        const checked = document.getElementById('selectAllOutbounds').checked;
        const deptFilter = document.getElementById('outboundDeptFilter').value;
        const visible = deptFilter ? Store.outbounds.filter(o => o.department === deptFilter) : Store.outbounds;
        if (checked) {
            visible.forEach(o => this.selectedIds.add(o.id));
        } else {
            visible.forEach(o => this.selectedIds.delete(o.id));
        }
        this.render();
    },

    async delete(id) {
        if (!confirm('确定删除该领用单？库存将回退。')) return;
        try {
            await API.delete(`outbounds/${id}`);
            this.selectedIds.delete(id);
            await Store.refresh(['products', 'outbounds']);
            this.render();
            Dashboard.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    },

    async batchDelete() {
        if (this.selectedIds.size === 0) {
            Utils.toast('请先勾选要删除的领用单', 'error');
            return;
        }
        const selected = Store.outbounds.filter(o => this.selectedIds.has(o.id));
        const preview = selected.slice(0, 5).map(o => o.no).join('、');
        const suffix = selected.length > 5 ? ` 等 ${selected.length} 张` : '';
        if (!confirm(`确定删除 ${preview}${suffix}？库存将回退。`)) return;

        try {
            for (const order of selected) {
                await API.delete(`outbounds/${order.id}`);
            }
            this.selectedIds.clear();
            await Store.refresh(['products', 'outbounds']);
            this.render();
            Dashboard.render();
            Utils.toast(`已删除 ${selected.length} 张领用单`);
        } catch (e) {
            Utils.toast(e.message, 'error');
            await Store.refresh(['products', 'outbounds']);
            this.render();
        }
    }
};

// ==================== 领用单模态框 ====================
const OutboundModal = {
    items: [],
    saving: false,

    open() {
        this.reset();
        document.getElementById('outboundModalTitle').textContent = '新建领用单';
        document.getElementById('outboundId').value = '';
        document.getElementById('outboundDate').value = Utils.today();
        Modals.open('outboundModal');
    },

    edit(id) {
        const o = Store.find('outbounds', id);
        if (!o) return;
        this.reset();
        document.getElementById('outboundModalTitle').textContent = '编辑领用单';
        document.getElementById('outboundId').value = o.id;
        document.getElementById('outboundDate').value = o.date;
        document.getElementById('outboundDept').value = o.department;
        document.getElementById('outboundPerson').value = o.person || '';

        this.items = o.items.map(i => ({
            productId: i.product_id,
            quantity: i.quantity,
            price: i.price,
            productName: i.product_name,
            productSpec: i.product_spec,
            productUnit: i.product_unit
        }));
        this.renderItems();
        Modals.open('outboundModal');
    },

    reset() {
        this.items = [];
        document.getElementById('outboundForm').reset();
        this.renderItems();
    },

    renderItems() {
        const tbody = document.getElementById('outboundItemsBody');
        if (this.items.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px;">点击"添加物品"按钮添加领用物品</td></tr>';
        } else {
            tbody.innerHTML = this.items.map((item, idx) => `
                <tr>
                    <td><strong>${item.productName || '未知'}</strong></td>
                    <td>${item.productSpec || '-'}</td>
                    <td>${item.productUnit || '-'}</td>
                    <td><input type="number" step="0.01" min="0.01" value="${item.quantity}" onchange="OutboundModal.updateItem(${idx}, 'quantity', this.value)" style="width:70px;"></td>
                    <td><input type="number" step="0.01" min="0" value="${item.price}" onchange="OutboundModal.updateItem(${idx}, 'price', this.value)" style="width:80px;"></td>
                    <td>${Utils.formatMoney(item.quantity * item.price)}</td>
                    <td><button type="button" class="btn btn-sm btn-danger" onclick="OutboundModal.removeItem(${idx})">删除</button></td>
                </tr>
            `).join('');
        }
        const total = this.items.reduce((sum, i) => sum + i.quantity * i.price, 0);
        document.getElementById('outboundTotal').textContent = Utils.formatMoney(total);
    },

    updateItem(idx, field, value) {
        if (field === 'productUnit') {
            this.items[idx][field] = String(value || '').trim();
        } else {
            this.items[idx][field] = parseFloat(value) || 0;
        }
        this.renderItems();
    },

    removeItem(idx) {
        this.items.splice(idx, 1);
        this.renderItems();
    },

    onDeptChange() {
        // 部门改变时，如果已有物品，提示可能需要重新选择
        if (this.items.length > 0) {
            // 不强制清空，让用户自己决定
        }
    },

    showProductSelector() {
        const dept = document.getElementById('outboundDept').value;
        if (!dept) {
            Utils.toast('请先选择领用部门', 'error');
            return;
        }

        const container = document.getElementById('outboundProductSelectorList');
        document.getElementById('outboundProductSearch').value = '';

        // 筛选：库存>0 且 未被选中 且 部门匹配
        const addedIds = new Set(this.items.map(i => i.productId));
        this._availableProducts = Store.products.filter(p => {
            if (p.stock <= 0) return false;
            if (addedIds.has(p.id)) return false;
            const depts = p.departments || [];
            if (depts.length === 0) return true; // 通用物品
            return depts.includes(dept);
        });

        this.renderProductCards(this._availableProducts, dept);
        Modals.open('outboundProductSelector');
    },

    renderProductCards(productList, dept) {
        const container = document.getElementById('outboundProductSelectorList');
        if (productList.length === 0) {
            container.innerHTML = '<div style="text-align:center;color:#999;padding:40px;">没有可领用的物品（库存不足或不属于该部门）</div>';
            return;
        }

        // 分组显示
        const deptProducts = productList.filter(p => {
            const depts = p.departments || [];
            return depts.includes(dept);
        });
        const generalProducts = productList.filter(p => {
            const depts = p.departments || [];
            return depts.length === 0;
        });

        let html = '';
        if (deptProducts.length > 0) {
            html += `<div style="font-size:13px;color:#1976d2;font-weight:600;margin:8px 0;">${dept}专属物品</div>`;
            html += deptProducts.map(p => this._productCardHtml(p)).join('');
        }
        if (generalProducts.length > 0) {
            html += `<div style="font-size:13px;color:#666;font-weight:600;margin:8px 0;">通用物品</div>`;
            html += generalProducts.map(p => this._productCardHtml(p)).join('');
        }
        container.innerHTML = html;
    },

    _productCardHtml(p) {
        return `
            <div class="product-select-card" data-id="${p.id}" onclick="this.classList.toggle('selected')">
                <div style="font-weight:500;color:#333;">${p.name} <span style="font-size:11px;color:#888;">(${p.code})</span></div>
                <div style="font-size:12px;color:#888;margin-top:4px;">${p.spec || '-'} | ${p.unit} | 库存 ${p.stock} | 单价 ¥${(p.cost || 0).toFixed(2)}</div>
                <div style="margin-top:8px;">
                    <label style="font-size:12px;color:#666;">数量：</label>
                    <input type="number" class="qty-input" step="0.01" min="0.01" value="1" style="width:70px;padding:4px;border:1px solid #ddd;border-radius:4px;" onclick="event.stopPropagation()">
                </div>
            </div>
        `;
    },

    filterProducts() {
        const search = document.getElementById('outboundProductSearch').value.toLowerCase().trim();
        const dept = document.getElementById('outboundDept').value;
        if (!search) {
            this.renderProductCards(this._availableProducts, dept);
            return;
        }
        const filtered = this._availableProducts.filter(p =>
            p.name.toLowerCase().includes(search) ||
            p.code.toLowerCase().includes(search)
        );
        this.renderProductCards(filtered, dept);
    },

    confirmProductSelect() {
        const selected = document.querySelectorAll('#outboundProductSelectorList .product-select-card.selected');
        selected.forEach(card => {
            const productId = card.dataset.id;
            const qty = parseFloat(card.querySelector('.qty-input').value) || 1;
            const product = Store.find('products', productId);
            if (product) {
                this.items.push({
                    productId: productId,
                    quantity: qty,
                    price: product.cost || 0,
                    productName: product.name,
                    productSpec: product.spec,
                    productUnit: product.unit
                });
            }
        });
        Modals.close('outboundProductSelector');
        this.renderItems();
    },

    async save(e) {
        e.preventDefault();
        if (this.saving) return;
        if (this.items.length === 0) {
            Utils.toast('请至少添加一个物品', 'error');
            return;
        }

        const saveBtn = document.getElementById('outboundSaveBtn');
        this.saving = true;
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.textContent = '保存中...';
        }

        const data = {
            id: document.getElementById('outboundId').value || undefined,
            department: document.getElementById('outboundDept').value,
            person: document.getElementById('outboundPerson').value.trim(),
            date: document.getElementById('outboundDate').value,
            items: this.items.map(i => ({ productId: i.productId, quantity: i.quantity, price: i.price }))
        };

        try {
            const result = await API.post('outbounds', data);
            await Store.refresh(['products', 'outbounds']);
            Modals.close('outboundModal');
            Outbounds.render();
            Dashboard.render();
            const saved = Store.find('outbounds', result.id);
            const no = saved?.no ? `（${saved.no}）` : '';
            Utils.toast(data.id ? `领用单更新成功${no}` : `领用单保存成功${no}，请勿重复录入`);
        } catch (e) {
            Utils.toast(e.message, 'error');
        } finally {
            this.saving = false;
            if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.textContent = '保存出库';
            }
        }
    }
};

// ==================== 领用单查看 ====================
const OutboundDetailModal = {
    currentId: null,

    open(id) {
        const o = Store.find('outbounds', id);
        if (!o) return;
        this.currentId = id;
        document.getElementById('outboundDetailContent').innerHTML = this.renderHtml(o, true);
        Modals.open('outboundDetailModal');
    },

    renderHtml(o, showPrice = true) {
        const headers = showPrice
            ? '<th>序号</th><th>物品名称</th><th>规格</th><th>单位</th><th>数量</th><th>单价</th><th>小计</th>'
            : '<th>序号</th><th>物品名称</th><th>规格</th><th>单位</th><th>数量</th>';
        const rows = (o.items || []).map((item, idx) => `
            <tr>
                <td>${idx + 1}</td>
                <td>${item.product_name || '未知'}</td>
                <td>${item.product_spec || '-'}</td>
                <td>${item.product_unit || '-'}</td>
                <td>${item.quantity}</td>
                ${showPrice ? `<td>${Utils.formatMoney(item.price)}</td><td>${Utils.formatMoney(item.subtotal)}</td>` : ''}
            </tr>
        `).join('');
        return `
            <div style="display:flex;justify-content:space-between;gap:16px;margin-bottom:16px;font-size:14px;">
                <div><strong>单号：</strong>${o.no}</div>
                <div><strong>日期：</strong>${o.date}</div>
            </div>
            <div style="display:flex;justify-content:space-between;gap:16px;margin-bottom:16px;font-size:14px;">
                <div><strong>领用部门：</strong>${o.department}</div>
                <div><strong>领用人：</strong>${o.person || '-'}</div>
            </div>
            <table class="data-table">
                <thead><tr>${headers}</tr></thead>
                <tbody>${rows || '<tr><td colspan="7" class="empty-state">暂无明细</td></tr>'}</tbody>
            </table>
            ${showPrice ? `<div class="total-row">合计：<span>${Utils.formatMoney(o.total)}</span></div>` : ''}
        `;
    },

    edit() {
        if (!this.currentId) return;
        const id = this.currentId;
        Modals.close('outboundDetailModal');
        OutboundModal.edit(id);
    },

    print() {
        if (!this.currentId) return;
        Printer.printOutbound(this.currentId);
    }
};

// ==================== 销售出库 ====================
const Sales = {
    selectedIds: new Set(),

    async ensureDetail(id) {
        const current = Store.find('sales', id);
        if (current && Array.isArray(current.items)) return current;
        const detail = await API.get(`sales/${id}`);
        Store.upsert('sales', detail);
        return Store.find('sales', id);
    },

    groupsByDate() {
        const groups = new Map();
        Store.sales.forEach(sale => {
            const date = sale.date || '未填写日期';
            if (!groups.has(date)) {
                groups.set(date, {
                    date,
                    ids: [],
                    customers: new Set(),
                    orderCount: 0,
                    itemCount: 0,
                    total: 0
                });
            }
            const group = groups.get(date);
            group.ids.push(sale.id);
            group.orderCount += 1;
            group.itemCount += Number(sale.item_count ?? (sale.items ? sale.items.length : 0));
            group.total += Number(sale.total || 0);
            if (sale.customer) group.customers.add(sale.customer);
        });
        return [...groups.values()].sort((a, b) => String(b.date).localeCompare(String(a.date)));
    },

    render() {
        const tbody = document.getElementById('salesList');
        if (Store.sales.length === 0) {
            this.selectedIds.clear();
            const selectAll = document.getElementById('salesSelectAll');
            if (selectAll) selectAll.checked = false;
            tbody.innerHTML = '<tr><td colspan="7" class="empty-state">暂无销售记录</td></tr>';
            return;
        }
        this.selectedIds = new Set([...this.selectedIds].filter(id => Store.sales.some(s => s.id === id)));
        const sales = [...Store.sales].sort((a, b) => {
            const dateCompare = String(b.date || '').localeCompare(String(a.date || ''));
            if (dateCompare !== 0) return dateCompare;
            return String(b.created_at || '').localeCompare(String(a.created_at || ''));
        });
        const selectAll = document.getElementById('salesSelectAll');
        if (selectAll) selectAll.checked = Store.sales.length > 0 && this.selectedIds.size === Store.sales.length;
        tbody.innerHTML = sales.map(sale => {
            const checked = this.selectedIds.has(sale.id);
            const itemCount = Number(sale.item_count ?? (sale.items ? sale.items.length : 0));
            return `
            <tr>
                <td class="sales-select-cell" data-label="选择"><input type="checkbox" class="sales-select" ${checked ? 'checked' : ''} onchange="Sales.toggleSelect('${sale.id}', this.checked)"></td>
                <td data-label="日期">${sale.date || '-'}</td>
                <td data-label="客户">${sale.customer || '-'}</td>
                <td data-label="单号">${sale.no || '-'}</td>
                <td data-label="物品数">${itemCount}</td>
                <td data-label="金额">${Utils.formatMoney(sale.total || 0)}</td>
                <td class="sales-row-actions" data-label="操作">
                    <button class="btn btn-sm" onclick="Printer.viewSales('${sale.id}')">查看</button>
                    <button class="btn btn-sm" onclick="SalesModal.edit('${sale.id}')">编辑</button>
                    <button class="btn btn-sm" onclick="Sales.sendPrintJob(['${sale.id}'], 'sales')">打印</button>
                    <button class="btn btn-sm" onclick="Sales.exportSale('${sale.id}')">下载销售单</button>
                    <button class="btn btn-sm" onclick="Sales.sendPrintJob(['${sale.id}'], 'purchase')">打印采购单</button>
                    <button class="btn btn-sm" onclick="Sales.exportPurchaseList('${sale.id}')">下载采购单</button>
                    <button class="btn btn-sm btn-danger" onclick="Sales.delete('${sale.id}')">删除</button>
                </td>
            </tr>
        `;
        }).join('');
    },

    toggleSelect(id, checked) {
        if (checked) {
            this.selectedIds.add(id);
        } else {
            this.selectedIds.delete(id);
        }
        const selectAll = document.getElementById('salesSelectAll');
        if (selectAll) selectAll.checked = Store.sales.length > 0 && this.selectedIds.size === Store.sales.length;
    },

    toggleDate(date, checked) {
        Store.sales.filter(s => (s.date || '未填写日期') === date).forEach(s => {
            if (checked) {
                this.selectedIds.add(s.id);
            } else {
                this.selectedIds.delete(s.id);
            }
        });
        const selectAll = document.getElementById('salesSelectAll');
        if (selectAll) selectAll.checked = Store.sales.length > 0 && this.selectedIds.size === Store.sales.length;
    },

    toggleSelectAll(checked) {
        this.selectedIds = checked ? new Set(Store.sales.map(s => s.id)) : new Set();
        document.querySelectorAll('.sales-select').forEach(input => input.checked = checked);
    },

    async editDate(date) {
        const sales = Store.sales.filter(s => (s.date || '未填写日期') === date);
        if (sales.length === 1) {
            await SalesModal.edit(sales[0].id);
            return;
        }
        Utils.toast('同一天有多张销售单，请在查看页面确认后单独编辑', 'error');
    },

    async exportSales(ids, paper = 'half', options = {}) {
        const saleIds = Array.isArray(ids) ? ids : [ids];
        if (saleIds.length === 0 || !saleIds[0]) {
            Utils.toast('请先选择要导出的销售单', 'error');
            return;
        }

        try {
            const res = await fetch(`${API_BASE}/api/sales/export`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ids: saleIds, paper, groupByCustomer: !!options.groupByCustomer })
            });
            if (!res.ok) {
                let message = `HTTP ${res.status}`;
                try {
                    const result = await res.json();
                    message = result.error || message;
                } catch (_) {}
                throw new Error(message);
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = options.filename || `销售单_${paper === 'third' ? '三等分' : '二等分'}_${Utils.today()}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
            Utils.toast(options.successMessage || `已导出 ${saleIds.length} 张${paper === 'third' ? '三等分' : '二等分'}销售单`);
        } catch (e) {
            Utils.toast('导出失败：' + e.message, 'error');
        }
    },

    exportSale(id, paper = 'half') {
        const sale = Store.sales.find(s => s.id === id);
        const name = sale ? `${sale.customer || '销售单'}_${sale.date || Utils.today()}` : `销售单_${Utils.today()}`;
        this.exportSales([id], paper, {
            filename: `${name}.xlsx`,
            successMessage: '已导出销售单'
        });
    },

    async exportSelected(paper = 'half') {
        const selected = Store.sales.filter(s => this.selectedIds.has(s.id));
        if (selected.length === 0) {
            Utils.toast('请先勾选要导出的销售单', 'error');
            return;
        }
        await this.exportSales(selected.map(s => s.id), paper, {
            groupByCustomer: true,
            filename: `销售单_按客户合并_${paper === 'third' ? '三等分' : '二等分'}_${Utils.today()}.xlsx`,
            successMessage: `已按客户合并导出 ${selected.length} 张销售单`
        });
    },

    async sendPrintJob(ids, type = 'sales', paper = 'half') {
        const saleIds = Array.isArray(ids) ? ids.filter(Boolean) : [ids].filter(Boolean);
        if (saleIds.length === 0) {
            Utils.toast('请先选择要打印的销售单', 'error');
            return;
        }
        try {
            const result = await API.post('print-jobs', { ids: saleIds, type, paper });
            const label = type === 'purchase' ? '采购单' : '销售单';
            Utils.toast(result.message || `已发送${label}到台式机打印队列`);
        } catch (e) {
            Utils.toast('发送打印失败：' + e.message, 'error');
        }
    },

    printSelected() {
        const selected = Store.sales.filter(s => this.selectedIds.has(s.id));
        if (selected.length === 0) {
            Utils.toast('请先勾选要打印的销售单', 'error');
            return;
        }
        this.sendPrintJob(selected.map(s => s.id), 'sales', 'half');
    },

    printSelectedPurchaseList() {
        const selected = Store.sales.filter(s => this.selectedIds.has(s.id));
        if (selected.length === 0) {
            Utils.toast('请先勾选要打印采购单的销售单', 'error');
            return;
        }
        this.sendPrintJob(selected.map(s => s.id), 'purchase', 'a4');
    },

    printTodayJobs() {
        const today = Utils.today();
        const deliverySales = Store.sales.filter(s => s.date === today);
        if (deliverySales.length === 0) {
            Utils.toast(`今天 ${today} 还没有销售单`, 'error');
            return;
        }
        const ids = deliverySales.map(s => s.id);
        this.sendPrintJob(ids, 'sales', 'half');
        this.sendPrintJob(ids, 'purchase', 'a4');
    },

    async exportPurchaseList(ids, filename = `采购单_${Utils.today()}.xlsx`) {
        const saleIds = Array.isArray(ids) ? ids : [ids];
        if (saleIds.length === 0 || !saleIds[0]) {
            Utils.toast('请先选择销售单', 'error');
            return;
        }

        try {
            const res = await fetch(`${API_BASE}/api/sales/purchase-list`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ids: saleIds })
            });
            if (!res.ok) {
                let message = `HTTP ${res.status}`;
                try {
                    const result = await res.json();
                    message = result.error || message;
                } catch (_) {}
                throw new Error(message);
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
            URL.revokeObjectURL(url);
            Utils.toast(saleIds.length === 1 ? '已导出采购单' : `已合并导出 ${saleIds.length} 张销售单的采购单`);
        } catch (e) {
            Utils.toast('导出采购单失败：' + e.message, 'error');
        }
    },

    async exportSelectedPurchaseList() {
        const selected = Store.sales.filter(s => this.selectedIds.has(s.id));
        if (selected.length === 0) {
            Utils.toast('请先勾选要合并采购的销售单', 'error');
            return;
        }
        await this.exportPurchaseList(selected.map(s => s.id), `合并采购单_${Utils.today()}.xlsx`);
    },

    async batchDelete() {
        const selected = Store.sales.filter(s => this.selectedIds.has(s.id));
        if (selected.length === 0) {
            Utils.toast('请先勾选要删除的销售单', 'error');
            return;
        }
        const preview = selected.slice(0, 5).map(s => s.no).join('、');
        const suffix = selected.length > 5 ? ` 等 ${selected.length} 张` : '';
        if (!confirm(`确定删除 ${preview}${suffix}？库存将按规则回退。`)) return;

        try {
            for (const sale of selected) {
                await API.delete(`sales/${sale.id}`);
            }
            this.selectedIds.clear();
            await Store.refresh(['products', 'sales']);
            this.render();
            Dashboard.render();
            Utils.toast(`已删除 ${selected.length} 张销售单`);
        } catch (e) {
            Utils.toast(e.message, 'error');
            await Store.refresh(['products', 'sales']);
            this.render();
        }
    },

    async delete(id) {
        if (!confirm('确定删除该销售单？库存将回退。')) return;
        try {
            await API.delete(`sales/${id}`);
            await Store.refresh(['products', 'sales']);
            this.selectedIds.delete(id);
            this.render();
            Dashboard.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    },

    async deleteDate(date) {
        const sales = Store.sales.filter(s => (s.date || '未填写日期') === date);
        if (sales.length === 0) return;
        if (!confirm(`确定删除 ${date} 的 ${sales.length} 张销售单？库存将按规则回退。`)) return;
        try {
            for (const sale of sales) {
                await API.delete(`sales/${sale.id}`);
            }
            sales.forEach(s => this.selectedIds.delete(s.id));
            await Store.refresh(['products', 'sales']);
            this.render();
            Dashboard.render();
            Utils.toast(`已删除 ${date} 的 ${sales.length} 张销售单`);
        } catch (e) {
            Utils.toast(e.message, 'error');
            await Store.refresh(['products', 'sales']);
            this.render();
        }
    }
};

const SalesDetailModal = {
    currentDate: null,

    async openDate(date) {
        this.currentDate = date;
        const sales = Store.sales.filter(s => (s.date || '未填写日期') === date);
        if (sales.length === 0) {
            Utils.toast('没有找到这一天的销售单', 'error');
            return;
        }
        const detailed = await Promise.all(sales.map(sale => Sales.ensureDetail(sale.id)));
        document.getElementById('salesDetailContent').innerHTML = this.renderDate(date, detailed);
        Modals.open('salesDetailModal');
    },

    renderDate(date, sales) {
        const allItems = [];
        sales.forEach(sale => {
            (sale.items || []).forEach(item => allItems.push({ sale, item }));
        });
        const customers = [...new Set(sales.map(s => s.customer).filter(Boolean))];
        const phones = [...new Set(sales.map(s => s.phone).filter(Boolean))];
        const orderNos = sales.map(s => s.no).filter(Boolean).join('、');
        const total = sales.reduce((sum, s) => sum + Number(s.total || 0), 0);
        const rows = allItems.map((row, idx) => {
            const item = row.item;
            return `
                <tr>
                    <td>${idx + 1}</td>
                    <td>${item.product_name || '未知'}</td>
                    <td>${item.product_spec || '-'}</td>
                    <td>${item.product_unit || '-'}</td>
                    <td>${item.quantity}</td>
                    <td>${Utils.formatMoney(item.price || 0)}</td>
                    <td>${Utils.formatMoney(item.subtotal || 0)}</td>
                </tr>
            `;
        }).join('');

        return `
            <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:14px;">
                <div><strong>日期：</strong>${date}</div>
                <div><strong>销售单数：</strong>${sales.length}</div>
                <div><strong>总金额：</strong>${Utils.formatMoney(total)}</div>
            </div>
            <div style="margin-bottom:10px;"><strong>客户：</strong>${customers.join('、') || '-'}</div>
            <div style="margin-bottom:10px;"><strong>电话：</strong>${phones.join('、') || '-'}</div>
            <div style="margin-bottom:16px;color:#666;font-size:13px;"><strong>销售单号：</strong>${orderNos || '-'}</div>
            <table class="data-table">
                <thead>
                    <tr><th>序号</th><th>物品名称</th><th>规格</th><th>单位</th><th>数量</th><th>单价</th><th>小计</th></tr>
                </thead>
                <tbody>${rows || '<tr><td colspan="7" class="empty-state">暂无物品</td></tr>'}</tbody>
            </table>
        `;
    },

    async downloadPurchaseList() {
        if (!this.currentDate) return;
        try {
            const res = await fetch(`${API_BASE}/api/sales/purchase-list?date=${encodeURIComponent(this.currentDate)}`);
            if (!res.ok) {
                let message = `HTTP ${res.status}`;
                try {
                    const result = await res.json();
                    message = result.error || message;
                } catch (_) {
                    const text = await res.text();
                    message = text ? text.slice(0, 80) : message;
                }
                throw new Error(message);
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `采购单_${this.currentDate}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
            Utils.toast('采购单已下载');
        } catch (e) {
            Utils.toast('下载采购单失败：' + e.message, 'error');
        }
    },

    print() {
        if (!this.currentDate) return;
        Printer.printSalesDate(this.currentDate);
    },

    edit() {
        if (!this.currentDate) return;
        const date = this.currentDate;
        Modals.close('salesDetailModal');
        Sales.editDate(date);
    }
};

// ==================== 销售单模态框 ====================
const SalesModal = {
    items: [],

    open() {
        this.reset();
        document.getElementById('salesModalTitle').textContent = '新建销售单';
        document.getElementById('salesId').value = '';
        document.getElementById('salesDate').value = Utils.today();
        this.renderCustomerOptions();
        Modals.open('salesModal');
    },

    async edit(id) {
        const s = await Sales.ensureDetail(id);
        if (!s) return;
        this.reset();
        document.getElementById('salesModalTitle').textContent = '编辑销售单';
        document.getElementById('salesId').value = s.id;
        document.getElementById('salesDate').value = s.date;
        document.getElementById('salesCompany').value = s.company || '';
        document.getElementById('salesCustomer').value = s.customer || '';
        document.getElementById('salesPhone').value = s.phone || '';
        document.getElementById('salesShowHandlers').checked = s.show_handlers === 1;
        document.getElementById('salesHandler').value = s.handler || '';
        document.getElementById('salesIssuer').value = s.issuer || '';
        document.getElementById('salesHandlerInputs').style.display = s.show_handlers === 1 ? 'flex' : 'none';
        this.renderCustomerOptions();

        this.items = s.items.map(i => ({
            productId: i.product_id,
            quantity: i.quantity,
            price: i.price,
            productName: i.product_name,
            productSpec: i.product_spec,
            productUnit: i.product_unit
        }));
        this.renderItems();
        Modals.open('salesModal');
    },

    reset() {
        this.items = [];
        document.getElementById('salesForm').reset();
        document.getElementById('salesHandlerInputs').style.display = 'none';
        this.renderItems();
    },

    renderCustomerOptions() {
        const select = document.getElementById('salesCustomerSelect');
        select.innerHTML = '<option value="">选择客户</option>' +
            Store.customers.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
    },

    onCustomerChange() {
        const customerId = document.getElementById('salesCustomerSelect').value;
        if (!customerId) return;
        const customer = Store.find('customers', customerId);
        if (customer) {
            document.getElementById('salesCompany').value = customer.company || '';
            document.getElementById('salesCustomer').value = customer.name;
            document.getElementById('salesPhone').value = customer.phone || '';
        }
    },

    toggleHandlers() {
        const checked = document.getElementById('salesShowHandlers').checked;
        document.getElementById('salesHandlerInputs').style.display = checked ? 'flex' : 'none';
    },

    renderItems() {
        const tbody = document.getElementById('salesItemsBody');
        if (this.items.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px;">点击"添加物品"按钮选择销售物品</td></tr>';
        } else {
            tbody.innerHTML = this.items.map((item, idx) => `
                <tr>
                    <td><strong>${item.productName || '未知'}</strong></td>
                    <td>${item.productSpec || '-'}</td>
                    <td><input type="text" value="${item.productUnit || ''}" onchange="SalesModal.updateItem(${idx}, 'productUnit', this.value)" style="width:58px;"></td>
                    <td><input type="number" step="0.01" min="0.01" value="${item.quantity}" onchange="SalesModal.updateItem(${idx}, 'quantity', this.value)" style="width:70px;"></td>
                    <td><input type="number" step="0.01" min="0" value="${item.price}" onchange="SalesModal.updateItem(${idx}, 'price', this.value)" style="width:80px;"></td>
                    <td>${Utils.formatMoney(item.quantity * item.price)}</td>
                    <td><button type="button" class="btn btn-sm btn-danger" onclick="SalesModal.removeItem(${idx})">删除</button></td>
                </tr>
            `).join('');
        }
        const total = this.items.reduce((sum, i) => sum + i.quantity * i.price, 0);
        document.getElementById('salesTotal').textContent = Utils.formatMoney(total);
    },

    updateItem(idx, field, value) {
        if (field === 'productUnit') {
            this.items[idx][field] = String(value || '').trim();
        } else {
            this.items[idx][field] = parseFloat(value) || 0;
        }
        this.renderItems();
    },

    removeItem(idx) {
        this.items.splice(idx, 1);
        this.renderItems();
    },

    showProductSelector() {
        const container = document.getElementById('salesProductSelectorList');
        document.getElementById('salesProductSearch').value = '';
        const addedIds = new Set(this.items.map(i => i.productId));
        this._availableProducts = Store.products.filter(p => !addedIds.has(p.id));
        this.renderProductCards(this._availableProducts);
        Modals.open('salesProductSelector');
    },

    renderProductCards(productList) {
        const container = document.getElementById('salesProductSelectorList');
        if (productList.length === 0) {
            container.innerHTML = '<div style="text-align:center;color:#999;padding:40px;">没有可添加的物品</div>';
            return;
        }
        container.innerHTML = productList.map(p => `
            <div class="product-select-card" data-id="${p.id}" onclick="this.classList.toggle('selected')">
                <div style="font-weight:500;color:#333;">${p.name}</div>
                <div style="font-size:12px;color:#888;margin-top:4px;">${p.spec || '-'} | ${p.unit} | 库存 ${p.stock} | 当前单价 ¥${(p.cost || 0).toFixed(2)}</div>
                <div style="margin-top:8px;">
                    <label style="font-size:12px;color:#666;">数量：</label>
                    <input type="number" class="qty-input" step="0.01" min="0.01" value="1" style="width:70px;padding:4px;border:1px solid #ddd;border-radius:4px;" onclick="event.stopPropagation()">
                </div>
            </div>
        `).join('');
    },

    filterProducts() {
        const search = document.getElementById('salesProductSearch').value.toLowerCase().trim();
        if (!search) {
            this.renderProductCards(this._availableProducts);
            return;
        }
        const filtered = this._availableProducts.filter(p =>
            p.name.toLowerCase().includes(search) ||
            p.code.toLowerCase().includes(search)
        );
        this.renderProductCards(filtered);
    },

    confirmProductSelect() {
        const selected = document.querySelectorAll('#salesProductSelectorList .product-select-card.selected');
        selected.forEach(card => {
            const productId = card.dataset.id;
            const qty = parseFloat(card.querySelector('.qty-input').value) || 1;
            const product = Store.find('products', productId);
            if (product) {
                this.items.push({
                    productId: productId,
                    quantity: qty,
                    price: product.last_sale_price || product.cost || 0,
                    productName: product.name,
                    productSpec: product.spec,
                    productUnit: product.unit
                });
            }
        });
        Modals.close('salesProductSelector');
        this.renderItems();
    },

    async quickAddProduct() {
        const name = prompt('输入新物品名称');
        if (!name || !name.trim()) return;
        const unit = prompt('输入单位，例如：斤、件、包、袋', '斤');
        if (!unit || !unit.trim()) return;
        const qtyText = prompt('本次数量', '1');
        const quantity = parseFloat(qtyText);
        if (!quantity || quantity <= 0) {
            Utils.toast('数量不正确', 'error');
            return;
        }
        const priceText = prompt('销售单价，不清楚可填 0', '0');
        const price = parseFloat(priceText) || 0;
        const code = Utils.generateNextCode('WP', Store.products);
        try {
            const result = await API.post('products', {
                code,
                name: name.trim(),
                spec: '',
                unit: unit.trim(),
                cost: 0,
                stock: 0,
                alertLine: 0,
                departments: [],
                lastSalePrice: price
            });
            await Store.refresh(['products']);
            const product = Store.find('products', result.id);
            this.items.push({
                productId: result.id,
                quantity,
                price,
                productName: product ? product.name : name.trim(),
                productSpec: '',
                productUnit: product ? product.unit : unit.trim()
            });
            this.renderItems();
            Utils.toast('已新增物品并加入销售单');
        } catch (e) {
            Utils.toast('新增物品失败：' + e.message, 'error');
        }
    },

    showPurchasePicker() {
        this.renderPurchasePickerList();
        Modals.open('purchasePickerModal');
    },

    renderPurchasePickerList() {
        const container = document.getElementById('purchasePickerList');
        const search = (document.getElementById('purchasePickerSearch')?.value || '').toLowerCase().trim();

        if (Store.purchases.length === 0) {
            container.innerHTML = '<div style="text-align:center;color:#999;padding:40px;">暂无采购单</div>';
            return;
        }

        let filtered = Store.purchases;
        if (search) {
            filtered = filtered.filter(p =>
                p.no.toLowerCase().includes(search) ||
                (p.supplier_name && p.supplier_name.toLowerCase().includes(search))
            );
        }

        if (filtered.length === 0) {
            container.innerHTML = '<div style="text-align:center;color:#999;padding:40px;">没有匹配的采购单</div>';
            return;
        }

        container.innerHTML = filtered.map(p => {
            const itemsHtml = p.items.map((it, idx) => `
                <div style="display:flex;align-items:center;gap:8px;padding:6px 10px;background:#f8f9fa;border-radius:4px;margin-top:4px;">
                    <input type="checkbox" class="pick-item-check" data-pid="${p.id}" data-productid="${it.product_id}" data-qty="${it.quantity}" data-price="${it.price}"
                        style="width:16px;height:16px;cursor:pointer;">
                    <span style="flex:1;font-size:13px;">${it.product_name || '未知'} (${it.product_id})</span>
                    <span style="font-size:12px;color:#888;">数量: ${it.quantity}</span>
                    <span style="font-size:12px;color:#888;">采购价: ¥${(it.price || 0).toFixed(2)}</span>
                </div>
            `).join('');

            return `
                <div style="border:1px solid #e9ecef;border-radius:8px;padding:12px;margin-bottom:12px;background:#fff;">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
                        <div>
                            <strong style="font-size:14px;color:#333;">${p.no}</strong>
                            <span style="color:#888;font-size:12px;margin-left:8px;">${p.date} | ${p.supplier_name || '-'}</span>
                        </div>
                        <label style="display:flex;align-items:center;gap:4px;cursor:pointer;font-size:12px;color:#666;">
                            <input type="checkbox" onchange="SalesModal.togglePickAll('${p.id}', this.checked)"> 全选
                        </label>
                    </div>
                    ${itemsHtml}
                </div>
            `;
        }).join('');
    },

    togglePickAll(purchaseId, checked) {
        document.querySelectorAll(`.pick-item-check[data-pid="${purchaseId}"]`).forEach(cb => cb.checked = checked);
    },

    filterPurchasePicker() {
        this.renderPurchasePickerList();
    },

    confirmPurchasePicker() {
        const checked = document.querySelectorAll('.pick-item-check:checked');
        if (checked.length === 0) {
            Utils.toast('请至少选择一个物品', 'error');
            return;
        }

        checked.forEach(cb => {
            const productId = cb.dataset.productid;
            const quantity = parseFloat(cb.dataset.qty) || 0;
            const product = Store.find('products', productId);
            const price = product && product.last_sale_price ? product.last_sale_price : (parseFloat(cb.dataset.price) || 0);

            const existing = this.items.find(i => i.productId === productId);
            if (existing) {
                existing.quantity += quantity;
            } else {
                this.items.push({
                    productId: productId,
                    quantity: quantity,
                    price: price,
                    productName: product ? product.name : '未知',
                    productSpec: product ? product.spec : '',
                    productUnit: product ? product.unit : ''
                });
            }
        });

        Modals.close('purchasePickerModal');
        this.renderItems();
    },

    async save(e) {
        e.preventDefault();
        await this.saveCurrent(false);
    },

    async saveAndPrint() {
        await this.saveCurrent(true);
    },

    async saveCurrent(sendPrint = false) {
        if (this.items.length === 0) {
            Utils.toast('请至少添加一个物品', 'error');
            return;
        }

        const data = {
            id: document.getElementById('salesId').value || undefined,
            company: document.getElementById('salesCompany').value.trim(),
            customer: document.getElementById('salesCustomer').value.trim(),
            phone: document.getElementById('salesPhone').value.trim(),
            date: document.getElementById('salesDate').value,
            items: this.items.map(i => ({
                productId: i.productId,
                quantity: i.quantity,
                price: i.price,
                productUnit: i.productUnit || ''
            })),
            showHandlers: document.getElementById('salesShowHandlers').checked,
            handler: document.getElementById('salesHandler').value.trim(),
            issuer: document.getElementById('salesIssuer').value.trim()
        };

        try {
            const result = await API.post('sales', data);
            await Store.refresh(['products', 'sales']);
            Modals.close('salesModal');
            Sales.render();
            Dashboard.render();
            if (sendPrint) {
                await Sales.sendPrintJob([result.id || data.id], 'sales', 'half');
            } else {
                Utils.toast(data.id ? '更新成功' : '保存成功');
            }
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 库存查询 ====================
const Inventory = {
    _requestSeq: 0,
    scheduleRender: null,

    async render() {
        const search = document.getElementById('inventorySearch').value;
        const status = document.getElementById('stockFilter').value;
        const requestSeq = ++this._requestSeq;

        try {
            const params = new URLSearchParams();
            if (search) params.append('search', search);
            if (status) params.append('status', status);

            const rows = await API.get(`inventory?${params.toString()}`);
            if (requestSeq !== this._requestSeq) return;
            const tbody = document.getElementById('inventoryList');

            if (rows.length === 0) {
                tbody.innerHTML = '<tr><td colspan="8" class="empty-state">暂无库存记录</td></tr>';
                return;
            }

            tbody.innerHTML = rows.map(p => {
                const statusClass = (p.status === 'normal' || p.status === 'flow') ? 'status-normal' : (p.status === 'low' ? 'status-warning' : 'status-danger');
                const statusText = p.status === 'flow' ? '生鲜流转' : (p.status === 'normal' ? '正常' : (p.status === 'low' ? '低库存' : '缺货'));
                return `
                    <tr>
                        <td>${p.code}</td>
                        <td><strong>${p.name}</strong></td>
                        <td>${p.spec || '-'}</td>
                        <td style="font-weight:600;${p.status !== 'normal' && p.status !== 'flow' ? 'color:#fa8c16;' : ''}">${p.stock}</td>
                        <td>${p.alert_line}</td>
                        <td>${p.cost > 0 ? Utils.formatMoney(p.cost) : '-'}</td>
                        <td>${Utils.formatMoney(p.stock * p.cost)}</td>
                        <td><span class="status ${statusClass}">${statusText}</span></td>
                    </tr>
                `;
            }).join('');
        } catch (e) {
            console.error('Inventory render error:', e);
        }
    },

    exportExcel() {
        window.location.href = '/api/export/inventory';
    }
};

Inventory.scheduleRender = Utils.debounce(() => Inventory.render(), 250);

// ==================== 部门统计 ====================
const DeptStats = {
    async render() {
        try {
            // 部门领用统计
            const deptData = await API.get('stats/departments');
            const stats = deptData.stats || [];
            const deptMap = { '食堂': 'deptCanteen', '保洁': 'deptClean', '办公': 'deptOffice', '维修': 'deptRepair', '酒店': 'deptHotel' };
            stats.forEach(s => {
                const el = document.getElementById(deptMap[s.department]);
                if (el) el.textContent = Utils.formatMoney(s.total);
            });

            // 部门领用明细
            const details = deptData.details || [];
            document.getElementById('deptDetailList').innerHTML = details.length === 0
                ? '<tr><td colspan="4" class="empty-state">暂无领用记录</td></tr>'
                : details.map(d => `
                    <tr>
                        <td>${d.department}</td>
                        <td>${d.order_count}</td>
                        <td>${d.item_count}</td>
                        <td>${Utils.formatMoney(d.total_amount)}</td>
                    </tr>
                `).join('');

            // 销售统计
            const salesData = await API.get('stats/sales');
            document.getElementById('salesOrderCount').textContent = salesData.orderCount;
            document.getElementById('salesTotalCost').textContent = Utils.formatMoney(salesData.totalCost);
            document.getElementById('salesTotalRevenue').textContent = Utils.formatMoney(salesData.totalRevenue);
            document.getElementById('salesTotalProfit').textContent = Utils.formatMoney(salesData.totalProfit);
            document.getElementById('salesAvgMargin').textContent = salesData.avgMargin.toFixed(1) + '%';

            // 销售明细
            const salesDetails = salesData.details || [];
            document.getElementById('salesDetailList').innerHTML = salesDetails.length === 0
                ? '<tr><td colspan="8" class="empty-state">暂无销售记录</td></tr>'
                : salesDetails.map(d => `
                    <tr>
                        <td>${d.customer}</td>
                        <td>${d.no}</td>
                        <td>${d.date}</td>
                        <td>${d.item_count}</td>
                        <td>${Utils.formatMoney(d.cost)}</td>
                        <td>${Utils.formatMoney(d.revenue)}</td>
                        <td style="color:${d.profit >= 0 ? '#52c41a' : '#f5222d'}">${Utils.formatMoney(d.profit)}</td>
                        <td>${d.margin.toFixed(1)}%</td>
                    </tr>
                `).join('');
        } catch (e) {
            console.error('DeptStats render error:', e);
        }
    },

    exportSales(mode) {
        window.location.href = `/api/export/sales-stats?mode=${mode}`;
    }
};

// ==================== 客户利润 ====================
const CustomerProfit = {
    data: null,
    selectedCustomer: '',
    month: '',

    async render() {
        try {
            const query = this.month ? `?month=${encodeURIComponent(this.month)}` : '';
            this.data = await API.get(`stats/customer-profits${query}`);
            const total = this.data.total || {};
            const revenue = total.revenue || 0;
            const profit = total.profit || 0;
            this.renderMonthOptions();
            document.getElementById('customerProfitCount').textContent = total.customerCount || 0;
            document.getElementById('customerProfitCost').textContent = Utils.formatMoney(total.cost || 0);
            document.getElementById('customerProfitRevenue').textContent = Utils.formatMoney(revenue);
            document.getElementById('customerProfitTotal').textContent = Utils.formatMoney(profit);
            document.getElementById('customerProfitMargin').textContent = (total.margin || 0).toFixed(1) + '%';

            const customers = this.data.customers || [];
            if (!customers.some(row => row.customer === this.selectedCustomer) && customers.length) {
                this.selectedCustomer = customers[0].customer;
            }
            document.getElementById('customerProfitList').innerHTML = customers.length === 0
                ? '<tr><td colspan="7" class="empty-state">暂无销售记录</td></tr>'
                : customers.map((row, idx) => `
                    <tr onclick="CustomerProfit.selectByIndex(${idx})" style="cursor:pointer;background:${row.customer === this.selectedCustomer ? '#f0f7ff' : ''}">
                        <td><strong>${row.customer}</strong></td>
                        <td>${row.orderCount}</td>
                        <td>${row.itemCount}</td>
                        <td>${Utils.formatMoney(row.cost)}</td>
                        <td>${Utils.formatMoney(row.revenue)}</td>
                        <td style="color:${row.profit >= 0 ? '#52c41a' : '#f5222d'}">${Utils.formatMoney(row.profit)}</td>
                        <td>${row.margin.toFixed(1)}%</td>
                    </tr>
                `).join('');
            this.renderCategories();
            this.renderItems();
            this.renderDetails();
        } catch (e) {
            console.error('CustomerProfit render error:', e);
        }
    },

    renderMonthOptions() {
        const select = document.getElementById('customerProfitMonth');
        const months = this.data.months || [];
        select.innerHTML = '<option value="">全部月份</option>' +
            months.map(month => `<option value="${month}" ${month === this.month ? 'selected' : ''}>${month}</option>`).join('');
    },

    changeMonth(month) {
        this.month = month;
        this.selectedCustomer = '';
        this.render();
    },

    selectByIndex(index) {
        const row = (this.data?.customers || [])[index];
        if (!row) return;
        this.selectedCustomer = row.customer;
        this.renderDetails();
        this.renderCustomerHighlights();
    },

    renderCustomerHighlights() {
        const customers = this.data?.customers || [];
        document.getElementById('customerProfitList').innerHTML = customers.length === 0
            ? '<tr><td colspan="7" class="empty-state">暂无销售记录</td></tr>'
            : customers.map((row, idx) => `
                <tr onclick="CustomerProfit.selectByIndex(${idx})" style="cursor:pointer;background:${row.customer === this.selectedCustomer ? '#f0f7ff' : ''}">
                    <td><strong>${row.customer}</strong></td>
                    <td>${row.orderCount}</td>
                    <td>${row.itemCount}</td>
                    <td>${Utils.formatMoney(row.cost)}</td>
                    <td>${Utils.formatMoney(row.revenue)}</td>
                    <td style="color:${row.profit >= 0 ? '#52c41a' : '#f5222d'}">${Utils.formatMoney(row.profit)}</td>
                    <td>${row.margin.toFixed(1)}%</td>
                </tr>
            `).join('');
    },

    renderCategories() {
        const categories = this.data?.categories || [];
        document.getElementById('customerProfitCategoryList').innerHTML = categories.length === 0
            ? '<tr><td colspan="7" class="empty-state">暂无品类数据</td></tr>'
            : categories.map(row => `
                <tr>
                    <td><strong>${row.category}</strong></td>
                    <td>${row.itemCount}</td>
                    <td>${Number(row.quantity || 0).toFixed(2)}</td>
                    <td>${Utils.formatMoney(row.cost)}</td>
                    <td>${Utils.formatMoney(row.revenue)}</td>
                    <td style="color:${row.profit >= 0 ? '#52c41a' : '#f5222d'}">${Utils.formatMoney(row.profit)}</td>
                    <td>${row.margin.toFixed(1)}%</td>
                </tr>
            `).join('');
    },

    renderItems() {
        const items = (this.data?.items || []).slice(0, 30);
        document.getElementById('customerProfitItemList').innerHTML = items.length === 0
            ? '<tr><td colspan="8" class="empty-state">暂无单品数据</td></tr>'
            : items.map(row => `
                <tr>
                    <td><strong>${row.name}</strong></td>
                    <td>${row.category}</td>
                    <td>${row.count}</td>
                    <td>${Number(row.quantity || 0).toFixed(2)} ${row.unit || ''}</td>
                    <td>${Utils.formatMoney(row.cost)}</td>
                    <td>${Utils.formatMoney(row.revenue)}</td>
                    <td style="color:${row.profit >= 0 ? '#52c41a' : '#f5222d'}">${Utils.formatMoney(row.profit)}</td>
                    <td>${row.margin.toFixed(1)}%</td>
                </tr>
            `).join('');
    },

    renderDetails() {
        const details = (this.data?.details || []).filter(row => (row.customer || '未填写客户') === this.selectedCustomer);
        document.getElementById('customerProfitDetailList').innerHTML = details.length === 0
            ? '<tr><td colspan="8" class="empty-state">请选择客户查看明细</td></tr>'
            : details.map(row => `
                <tr>
                    <td>${row.customer || '未填写客户'}</td>
                    <td>${row.no}</td>
                    <td>${row.date}</td>
                    <td>${row.item_count}</td>
                    <td>${Utils.formatMoney(row.cost)}</td>
                    <td>${Utils.formatMoney(row.revenue)}</td>
                    <td style="color:${row.profit >= 0 ? '#52c41a' : '#f5222d'}">${Utils.formatMoney(row.profit)}</td>
                    <td>${row.margin.toFixed(1)}%</td>
                </tr>
            `).join('');
    },

    exportReport() {
        const query = this.month ? `?month=${encodeURIComponent(this.month)}` : '';
        window.location.href = `/api/export/customer-profit-report${query}`;
    }
};

// ==================== 账单统计 ====================
const Finance = {
    async render() {
        try {
            const data = await API.get('statement');
            document.getElementById('totalPurchase').textContent = Utils.formatMoney(data.totalPurchase);
            document.getElementById('totalOutbound').textContent = Utils.formatMoney(data.totalOutbound);
            document.getElementById('totalOtherExpense').textContent = Utils.formatMoney(data.totalOtherExpense);

            const tbody = document.getElementById('financeList');
            const records = data.records || [];
            if (records.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" class="empty-state">暂无记录</td></tr>';
                return;
            }

            tbody.innerHTML = records.map(r => `
                <tr>
                    <td>${r.date}</td>
                    <td>${r.type}</td>
                    <td>${r.category || '-'}</td>
                    <td style="color:${r.type === '收入' ? '#52c41a' : '#f5222d'};font-weight:600;">${Utils.formatMoney(r.amount)}</td>
                    <td>${r.note || '-'}</td>
                    <td>
                        ${r.type === '财务记账' ? `<button class="btn btn-sm btn-danger" onclick="Finance.deleteRecord('${r.id}')">删除</button>` : '-'}
                    </td>
                </tr>
            `).join('');
        } catch (e) {
            console.error('Finance render error:', e);
        }
    },

    async deleteRecord(id) {
        if (!confirm('确定删除该记录？')) return;
        try {
            await API.delete(`finances/${id}`);
            await Store.refresh('finances');
            this.render();
            Utils.toast('删除成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    },

    exportStatement() {
        window.location.href = '/api/export/statement';
    }
};

// ==================== 财务记账模态框 ====================
const FinanceModal = {
    open() {
        document.getElementById('financeModalTitle').textContent = '记一笔';
        document.getElementById('financeForm').reset();
        document.getElementById('financeId').value = '';
        document.getElementById('financeDate').value = Utils.today();
        Modals.open('financeModal');
    },

    async save(e) {
        e.preventDefault();
        const data = {
            id: document.getElementById('financeId').value || undefined,
            date: document.getElementById('financeDate').value,
            type: document.getElementById('financeType').value,
            category: document.getElementById('financeCategory').value.trim(),
            amount: parseFloat(document.getElementById('financeAmount').value) || 0,
            note: document.getElementById('financeNote').value.trim()
        };

        try {
            await API.post('finances', data);
            await Store.refresh('finances');
            Modals.close('financeModal');
            Finance.render();
            Utils.toast('保存成功');
        } catch (e) {
            Utils.toast(e.message, 'error');
        }
    }
};

// ==================== 打印功能 ====================
const Printer = {
    defaultPrintSettings: {
        salesWidth: 241,
        salesHeight: 140,
        salesMarginTop: 3,
        salesMarginRight: 4,
        salesMarginBottom: 3,
        salesMarginLeft: 4,
        salesFontSize: 9.2,
        salesRowHeight: 3.9,
        purchaseMargin: 10,
        purchaseFontSize: 10.2
    },

    getPrintSettings() {
        try {
            const saved = JSON.parse(localStorage.getItem('hotelPrintSettings') || '{}');
            return { ...this.defaultPrintSettings, ...saved };
        } catch (_) {
            return { ...this.defaultPrintSettings };
        }
    },

    fillPrintSettingsForm() {
        const s = this.getPrintSettings();
        const fields = {
            printSalesWidth: 'salesWidth',
            printSalesHeight: 'salesHeight',
            printSalesMarginTop: 'salesMarginTop',
            printSalesMarginRight: 'salesMarginRight',
            printSalesMarginBottom: 'salesMarginBottom',
            printSalesMarginLeft: 'salesMarginLeft',
            printSalesFontSize: 'salesFontSize',
            printSalesRowHeight: 'salesRowHeight',
            printPurchaseMargin: 'purchaseMargin',
            printPurchaseFontSize: 'purchaseFontSize'
        };
        Object.entries(fields).forEach(([id, key]) => {
            const input = document.getElementById(id);
            if (input) input.value = s[key];
        });
    },

    openPrintSettings() {
        this.fillPrintSettingsForm();
        Modals.open('printSettingsModal');
    },

    savePrintSettings() {
        const read = (id, fallback) => {
            const value = parseFloat(document.getElementById(id)?.value);
            return Number.isFinite(value) && value > 0 ? value : fallback;
        };
        const defaults = this.defaultPrintSettings;
        const settings = {
            salesWidth: read('printSalesWidth', defaults.salesWidth),
            salesHeight: read('printSalesHeight', defaults.salesHeight),
            salesMarginTop: read('printSalesMarginTop', defaults.salesMarginTop),
            salesMarginRight: read('printSalesMarginRight', defaults.salesMarginRight),
            salesMarginBottom: read('printSalesMarginBottom', defaults.salesMarginBottom),
            salesMarginLeft: read('printSalesMarginLeft', defaults.salesMarginLeft),
            salesFontSize: read('printSalesFontSize', defaults.salesFontSize),
            salesRowHeight: read('printSalesRowHeight', defaults.salesRowHeight),
            purchaseMargin: read('printPurchaseMargin', defaults.purchaseMargin),
            purchaseFontSize: read('printPurchaseFontSize', defaults.purchaseFontSize)
        };
        localStorage.setItem('hotelPrintSettings', JSON.stringify(settings));
        Modals.close('printSettingsModal');
        Utils.toast('打印设置已保存');
    },

    resetPrintSettings() {
        localStorage.removeItem('hotelPrintSettings');
        this.fillPrintSettingsForm();
        Utils.toast('已恢复默认打印设置');
    },

    escapeHtml(value) {
        return String(value ?? '').replace(/[&<>"']/g, ch => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[ch]));
    },

    purchaseExportNote(name, notes) {
        const simpleNotes = new Set(['斤', '个', '件', '包', '袋', '箱', '瓶', '盒', '盘', '桶', '把', '根', '条', '颗', '捆', '提', '只', '块', '张', '份', '卷', '板']);
        const itemName = String(name || '');
        const noteAliases = { '连菜': '莲菜' };
        return [...notes]
            .map(note => String(note || '').trim())
            .filter(note => {
                const normalizedNote = noteAliases[note] || note;
                return note && !simpleNotes.has(note) && !itemName.includes(note) && normalizedNote !== itemName;
            })
            .filter((note, idx, arr) => arr.indexOf(note) === idx)
            .join('；');
    },

    purchaseCategory(name) {
        const text = String(name || '');
        const priorityGroups = [
            ['豆制品类', ['黄豆芽', '大豆芽', '小豆芽', '豆芽菜', '凉粉', '红豆腐丝', '黑豆腐丝', '素毛肚丝', '精品魔芋', '方块面筋', '方块儿面筋', '羊肚丝', '酸豆角']],
            ['冻货类', ['花卷']],
            ['粮油调料类', ['素鸡', '泡椒', '辣椒段', '辣皮子', '白糖']],
            ['其他类', ['六六红', '水晶粉', '火锅宽粉', '火锅邵皮', '火锅苕皮', '芝麻酱', '贡菜', '金针菇', '针金菇', '青岛纯生', '鞭炮笋', '香铃卷', '鲜玉米', '鲜虾']],
        ];
        for (const [label, keywords] of priorityGroups) {
            if (keywords.some(keyword => text.includes(keyword))) return label;
        }
        const groups = [
            ['蔬菜类', ['菜', '广东菜心', '连菜', '莲菜', '豆芽', '莲花白', '花白', '白菜', '青菜', '菠菜', '芹', '毛芹', '大芹菜', '蒜苔', '蒜薹', '蒜苗', '新蒜', '扒皮蒜', '去皮蒜', '净蒜', '豆王', '长江豆', '黄瓜', '西红柿', '青椒', '红椒', '红辣椒', '青辣椒', '土豆', '红薯', '胡萝卜', '白萝卜', '红心萝卜', '樱桃小萝卜', '萝卜苗', '紫甘蓝', '西葫芦', '南瓜', '冬瓜', '洋葱', '茄子', '平菇', '香菇', '蘑菇', '葱', '西兰花', '龙须菜', '甜豆', '广红']],
            ['水果类', ['西瓜', '哈密瓜', '火龙果', '柠檬', '油桃', '桃子', '苹果', '香蕉', '小香蕉', '圣女果', '乳瓜', '青提', '梨', '橙']],
            ['肉蛋类', ['肉', '鸡蛋', '鸭蛋', '咸鸭蛋', '猪耳朵']],
            ['饮品乳品类', ['牛奶', '酸奶', '蓝莓酱', '苹果酱']],
            ['主食面点类', ['面', '凉皮', '馒头', '包子', '油条', '饺子皮', '吐司', '法棍', '牛角包', '发糕']],
            ['冻货类', ['冻', '丸子', '馄饨', '鸡块']],
            ['豆制品类', ['豆腐', '豆干', '豆皮', '腐竹', '凉粉', '酸豆角', '红豆腐丝', '黑豆腐丝', '素毛肚丝', '羊肚丝', '方块面筋', '方块儿面筋', '面筋']],
            ['粮油调料类', ['大米', '江米', '油', '白醋', '生抽', '酱', '辣椒', '辣椒段', '辣皮子', '泡椒', '素鸡', '白糖', '粉条', '粉丝', '米线', '红九九']],
            ['饮品乳品类', ['冰峰', '酸梅汤', '果汁', '饼干']]
        ];
        for (const [label, keywords] of groups) {
            if (keywords.some(keyword => text.includes(keyword))) return label;
        }
        return '其他类';
    },

    purchaseSortRank(category) {
        const ranks = {
            '青菜供应商': 1,
            '王伟供应商': 2,
            '自采': 3,
            '蔬菜类': 10,
            '豆制品类': 11,
            '水果类': 12,
            '肉蛋类': 13,
            '主食面点类': 14,
            '粮油调料类': 15,
            '饮品乳品类': 16,
            '冻货类': 17,
            '饼干': 18,
            '其他类': 99
        };
        return ranks[category] || 99;
    },

    purchaseSupplierGroup(name) {
        const text = String(name || '').trim();
        const groups = {
            '青菜供应商': ['菠菜', '麦芹', '香菜', '圆生菜', '青菜（把）', '青菜', '小青菜', '广东菜心', '菜心', '香葱', '小葱'],
            '王伟供应商': ['豇豆', '长豇豆', '白豆角', '螺丝椒', '蒜苔', '蒜薹', '韭菜'],
            '自采': ['纯牛奶', '牛奶', '水果玉米', '青笋', '平菇', '黄瓜', '西兰花', '西蓝花', '红萝卜', '胡萝卜', '洋葱', '新蒜', '净蒜', '蒜苗', '白萝卜', '象牙萝卜', '生姜', '姜']
        };
        for (const [group, names] of Object.entries(groups)) {
            if (names.includes(text)) return group;
        }
        return '';
    },

    purchaseSupplierItemRank(name) {
        const order = [
            '菠菜', '麦芹', '香菜', '圆生菜', '青菜（把）', '青菜', '小青菜', '广东菜心', '菜心', '香葱', '小葱',
            '豇豆', '长豇豆', '螺丝椒', '蒜苔', '蒜薹', '韭菜', '白豆角',
            '纯牛奶', '牛奶', '水果玉米', '青笋', '平菇', '黄瓜', '西兰花', '西蓝花', '红萝卜', '胡萝卜', '洋葱', '新蒜', '净蒜', '蒜苗', '白萝卜', '象牙萝卜', '生姜', '姜'
        ];
        const index = order.indexOf(String(name || '').trim());
        return index >= 0 ? index : 9999;
    },

    purchaseGroupLabel(name) {
        return this.purchaseSupplierGroup(name) || this.purchaseCategory(name);
    },

    vegetableSortRank(name) {
        const order = [
            '青菜（把）', '青菜', '小青菜', '黄瓜',
            '水果玉米', '韭黄', '菜花', '花菜', '有机菜花', '红萝卜', '胡萝卜', '广红',
            '西红柿', '莲菜', '娃娃菜', '豆苗', '广东菜心',
            '小葱', '香葱', '香菜',
            '白菜', '大白菜', '罗马生菜', '生菜', '叶生菜', '圆生菜',
            '豇豆', '长豇豆', '长江豆', '长豆角',
            '韭菜', '芹菜', '麦芹', '毛芹',
            '白萝卜',
            '土豆', '黄土豆', '红薯', '蒜苔', '蒜薹',
            '贝贝南瓜', '南瓜', '西葫芦', '青椒',
            '平菇', '杏鲍菇',
            '包菜', '花白', '莲花白', '莲莲花白', '脆花白', '黄芽菜', '娃娃菜',
            '菜心', '奶白菜', '菠菜', '波菜', '芥菜', '苦菊',
            '西兰花', '西蓝花', '紫甘蓝',
            '山药', '水洗铁棍山药', '青笋', '净笋',
            '生姜', '净蒜', '蒜苗', '大葱', '白葱', '洋葱',
            '冬瓜', '苦瓜',
            '茄子', '广茄', '圆茄', '圆茄子',
            '青椒', '红椒', '青辣椒', '红辣椒', '线椒', '红线椒', '绿线椒', '螺丝椒',
            '白玉菇', '金针菇', '平菇', '蘑菇', '磨菇', '香菇', '杏鲍菇',
            '白豆角', '豆王', '荷兰豆', '豇豆', '长豆角', '长豇豆', '豆苗', '水果玉米',
        ];
        const index = order.indexOf(String(name || '').trim());
        return index >= 0 ? index : 9999;
    },

    purchaseNameSortRank(name) {
        const order = [
            '青菜（把）', '青菜', '小青菜', '小葱', '香葱', '香菜',
            '白菜', '大白菜', '罗马生菜',
            '豇豆', '长豇豆', '长江豆', '长豆角',
            '韭菜', '韭黄', '芹菜', '麦芹', '毛芹',
            '菜花', '花菜', '有机菜花', '白萝卜', '胡萝卜', '红萝卜', '广红',
            '土豆', '黄土豆', '红薯', '蒜苔', '蒜薹', '黄瓜',
            '贝贝南瓜', '南瓜', '西葫芦', '西红柿', '青椒',
            '平菇', '杏鲍菇', '水果玉米', '莲菜',
            '大豆芽', '豆芽', '黄豆芽', '小豆芽', '酸豆角', '凉粉',
            '圣女果', '桃子', '西瓜', '白心火龙果', '火龙果', '苹果', '香蕉', '哈密瓜', '青提', '油桃',
            '五花肉片', '排骨', '猪耳朵', '鸡蛋',
            '干米线', '细薄韭叶面', '肉包子', '臊子面', '馒头',
            '海带丝', '白糖',
            '冰红茶', '瓶装冰峰', '纯牛奶', '益生菌酸奶', '蓝莓酱', '苹果酱',
            '冻玉米', '火锅丸子', '里昂火腿', '馄饨', '花卷',
            '原味面包', '黑全麦面包', '豆浆粉',
            '六六红', '水晶粉', '火锅宽粉', '火锅邵皮', '火锅苕皮', '芝麻酱', '贡菜', '金针菇', '针金菇', '鞭炮笋', '香铃卷', '鲜玉米', '鲜虾', '青岛纯生',
        ];
        const index = order.indexOf(String(name || '').trim());
        return index >= 0 ? index : 9999;
    },

    sortedPurchaseItems(items) {
        return [...(items || [])].sort((a, b) => {
            const aName = a.name || a.product_name || '';
            const bName = b.name || b.product_name || '';
            const aCategory = a.category || this.purchaseGroupLabel(aName);
            const bCategory = b.category || this.purchaseGroupLabel(bName);
            const categoryRank = this.purchaseSortRank(aCategory) - this.purchaseSortRank(bCategory);
            if (categoryRank !== 0) return categoryRank;
            const supplierRank = this.purchaseSupplierItemRank(aName) - this.purchaseSupplierItemRank(bName);
            if (supplierRank !== 0) return supplierRank;
            const vegetableRank = (aCategory === '蔬菜类' ? this.vegetableSortRank(aName) : 9999) - (bCategory === '蔬菜类' ? this.vegetableSortRank(bName) : 9999);
            if (vegetableRank !== 0) return vegetableRank;
            const nameRank = this.purchaseNameSortRank(aName) - this.purchaseNameSortRank(bName);
            if (nameRank !== 0) return nameRank;
            return String(aName).localeCompare(String(bName), 'zh-Hans-CN');
        });
    },

    buildPurchaseRowsFromSales(ids) {
        const sales = ids.map(id => Store.find('sales', id)).filter(Boolean);
        const customers = [...new Set(sales.map(s => s.customer || '未填写客户'))];
        const grouped = new Map();
        sales.forEach(sale => {
            (sale.items || []).forEach(item => {
                const name = item.product_name || '未知';
                const unit = item.product_unit || '';
                const key = `${name}||${unit}`;
                if (!grouped.has(key)) {
                    grouped.set(key, {
                        name,
                        unit,
                        quantity: 0,
                        byCustomer: {},
                        notes: new Set(),
                        category: this.purchaseGroupLabel(name)
                    });
                }
                const row = grouped.get(key);
                const qty = Number(item.quantity || 0);
                const customer = sale.customer || '未填写客户';
                row.quantity += qty;
                row.byCustomer[customer] = (row.byCustomer[customer] || 0) + qty;
                if (item.note) row.notes.add(item.note);
            });
        });
        const rows = this.sortedPurchaseItems([...grouped.values()]);
        return { sales, customers, rows };
    },

    printPurchaseListFromSales(ids) {
        const { sales, customers, rows } = this.buildPurchaseRowsFromSales(ids);
        if (!sales.length || !rows.length) {
            Utils.toast('没有可打印的采购单明细', 'error');
            return;
        }
        const dates = [...new Set(sales.map(s => s.date).filter(Boolean))];
        const titleDate = dates.length === 1 ? dates[0] : Utils.today();
        const title = sales.length === 1
            ? `${sales[0].customer || '酒店'}采购单_${titleDate}`
            : `合并采购单_${titleDate}`;
        const showCustomerColumns = sales.length > 1 && customers.length > 1;
        let currentCategory = '';
        const customerHeaders = showCustomerColumns
            ? customers.map(customer => `<th>${this.escapeHtml(customer)}</th>`).join('')
            : '';
        const bodyRows = rows.map(row => {
            const category = row.category === currentCategory ? '' : row.category;
            currentCategory = row.category;
            const customerCells = showCustomerColumns
                ? customers.map(customer => `<td class="num">${row.byCustomer[customer] ? Utils.formatQty(row.byCustomer[customer]) : ''}</td>`).join('')
                : '';
            const noteCell = this.purchaseExportNote(row.name, row.notes);
            return `
                <tr>
                    <td>${this.escapeHtml(category)}</td>
                    <td>${this.escapeHtml(row.name)}</td>
                    <td class="num">${Utils.formatQty(row.quantity)}</td>
                    <td>${this.escapeHtml(row.unit)}</td>
                    ${customerCells}
                    <td>${this.escapeHtml(noteCell)}</td>
                </tr>
            `;
        }).join('');
        const html = `
            <section class="purchase-print-page">
                <h2>${this.escapeHtml(title)}</h2>
                <div class="purchase-meta">
                    <span>日期：${this.escapeHtml(titleDate)}</span>
                    <span>销售单数：${sales.length}</span>
                </div>
                <table class="purchase-print-table">
                    <thead>
                        <tr>
                            <th style="width:18%;">食材分类</th>
                            <th style="width:28%;">食材名称</th>
                            <th style="width:14%;">采购数量</th>
                            <th style="width:12%;">计量单位</th>
                            ${customerHeaders}
                            <th style="width:16%;">备注</th>
                        </tr>
                    </thead>
                    <tbody>${bodyRows}</tbody>
                </table>
                <div class="purchase-sign-row">
                    <span>采购人：__________</span>
                    <span>验收人：__________</span>
                    <span>日期：__________</span>
                </div>
            </section>
        `;
        this.openDocument(html, true, 'PURCHASE_A4');
    },

    printPurchase(id) {
        const p = Store.find('purchases', id);
        if (!p) return;

        const itemsHtml = p.items.map((item, idx) => `
            <tr>
                <td>${idx + 1}</td>
                <td>${item.product_name || '未知'}</td>
                <td>${item.product_spec || '-'}</td>
                <td>${item.product_unit || '-'}</td>
                <td>${item.quantity}</td>
                <td>¥${(item.price || 0).toFixed(2)}</td>
                <td>¥${(item.subtotal || 0).toFixed(2)}</td>
            </tr>
        `).join('');

        const html = `
            <div style="padding:20px;font-family:'Microsoft YaHei',sans-serif;">
                <h2 style="text-align:center;margin-bottom:20px;">采购入库单</h2>
                <div style="display:flex;justify-content:space-between;margin-bottom:20px;font-size:14px;">
                    <span>单号：${p.no}</span>
                    <span>日期：${p.date}</span>
                </div>
                <div style="margin-bottom:20px;font-size:14px;">供应商：${p.supplier_name || '-'}</div>
                <table style="width:100%;border-collapse:collapse;font-size:13px;">
                    <thead>
                        <tr style="background:#f5f5f5;">
                            <th style="border:1px solid #ddd;padding:8px;">序号</th>
                            <th style="border:1px solid #ddd;padding:8px;">物品名称</th>
                            <th style="border:1px solid #ddd;padding:8px;">规格</th>
                            <th style="border:1px solid #ddd;padding:8px;">单位</th>
                            <th style="border:1px solid #ddd;padding:8px;">数量</th>
                            <th style="border:1px solid #ddd;padding:8px;">单价</th>
                            <th style="border:1px solid #ddd;padding:8px;">小计</th>
                        </tr>
                    </thead>
                    <tbody>${itemsHtml}</tbody>
                </table>
                <div style="text-align:right;margin-top:16px;font-size:16px;font-weight:600;">
                    合计：${Utils.formatMoney(p.total)}
                </div>
                <div style="margin-top:40px;display:flex;justify-content:space-between;font-size:14px;">
                    <span>采购人：__________</span>
                    <span>验收人：__________</span>
                    <span>日期：__________</span>
                </div>
            </div>
        `;
        this.doPrint(html);
    },

    printOutbound(id) {
        const o = Store.find('outbounds', id);
        if (!o) return;
        const showPrice = document.getElementById('printShowPrice').checked;

        const itemsHtml = o.items.map((item, idx) => {
            if (showPrice) {
                return `
                    <tr>
                        <td style="border:1px solid #ddd;padding:8px;">${idx + 1}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.product_name || '未知'}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.product_spec || '-'}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.product_unit || '-'}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.quantity}</td>
                        <td style="border:1px solid #ddd;padding:8px;">¥${(item.price || 0).toFixed(2)}</td>
                        <td style="border:1px solid #ddd;padding:8px;">¥${(item.subtotal || 0).toFixed(2)}</td>
                    </tr>
                `;
            } else {
                return `
                    <tr>
                        <td style="border:1px solid #ddd;padding:8px;">${idx + 1}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.product_name || '未知'}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.product_spec || '-'}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.product_unit || '-'}</td>
                        <td style="border:1px solid #ddd;padding:8px;">${item.quantity}</td>
                    </tr>
                `;
            }
        }).join('');

        const totalRow = showPrice ? `
            <div style="text-align:right;margin-top:16px;font-size:16px;font-weight:600;">
                合计：${Utils.formatMoney(o.total)}
            </div>
        ` : '';

        const headers = showPrice
            ? '<th style="border:1px solid #ddd;padding:8px;">序号</th><th style="border:1px solid #ddd;padding:8px;">物品名称</th><th style="border:1px solid #ddd;padding:8px;">规格</th><th style="border:1px solid #ddd;padding:8px;">单位</th><th style="border:1px solid #ddd;padding:8px;">数量</th><th style="border:1px solid #ddd;padding:8px;">单价</th><th style="border:1px solid #ddd;padding:8px;">小计</th>'
            : '<th style="border:1px solid #ddd;padding:8px;">序号</th><th style="border:1px solid #ddd;padding:8px;">物品名称</th><th style="border:1px solid #ddd;padding:8px;">规格</th><th style="border:1px solid #ddd;padding:8px;">单位</th><th style="border:1px solid #ddd;padding:8px;">数量</th>';

        const html = `
            <div style="padding:20px;font-family:'Microsoft YaHei',sans-serif;">
                <h2 style="text-align:center;margin-bottom:20px;">部门领用单</h2>
                <div style="display:flex;justify-content:space-between;margin-bottom:20px;font-size:14px;">
                    <span>单号：${o.no}</span>
                    <span>日期：${o.date}</span>
                </div>
                <div style="margin-bottom:20px;font-size:14px;">领用部门：${o.department} | 领用人：${o.person || '-'}</div>
                <table style="width:100%;border-collapse:collapse;font-size:13px;">
                    <thead>
                        <tr style="background:#f5f5f5;">${headers}</tr>
                    </thead>
                    <tbody>${itemsHtml}</tbody>
                </table>
                ${totalRow}
                <div style="margin-top:40px;display:flex;justify-content:space-between;font-size:14px;">
                    <span>领用人签字：__________</span>
                    <span>审批人：__________</span>
                    <span>日期：__________</span>
                </div>
            </div>
        `;
        this.doPrint(html);
    },

    async viewSales(id) {
        await this.showSales(id, false);
    },

    viewSalesDate(date) {
        this.showSalesDate(date, false);
    },

    async printSales(id) {
        await this.showSales(id, true);
    },

    printSalesDate(date) {
        this.showSalesDate(date, true);
    },

    renderSalesDocument(s, startPageIndex = 0) {
        const company = s.company || '西安禾润佳商贸有限公司';
        const items = this.sortedPurchaseItems(s.items || []);
        const pages = [];
        for (let i = 0; i < items.length; i += 20) {
            pages.push(items.slice(i, i + 20));
        }
        if (pages.length === 0) pages.push([]);
        const handlerRow = s.show_handlers ? `
            <div class="sale-sign-row">
                <span>经手人：${s.handler || '__________'}</span>
                <span>出库人：${s.issuer || '__________'}</span>
            </div>
        ` : '';

        return pages.map((pageItems, pageIndex) => {
            const globalPageIndex = startPageIndex + pageIndex;
            const rowsHtml = pageItems.map((item, idx) => `
                <tr>
                    <td>${pageIndex * 20 + idx + 1}</td>
                    <td>${item.product_name || '未知'}</td>
                    <td>${item.product_spec || ''}</td>
                    <td class="num">${Number(item.quantity || 0).toFixed(Number.isInteger(Number(item.quantity || 0)) ? 0 : 2)}</td>
                    <td>${item.product_unit || '-'}</td>
                    <td class="num">${Number(item.price || 0).toFixed(2)}</td>
                    <td class="num">${Number(item.subtotal || 0).toFixed(2)}</td>
                    <td>${this.escapeHtml(item.note || '')}</td>
                </tr>
            `).join('');
            return `
                <section class="sales-print-page ${globalPageIndex > 0 ? 'page-break' : ''}">
                    <h2>${company}</h2>
                    <div class="sale-meta sale-meta-grid">
                        <span>客户：${s.customer || '-'}</span>
                        <span>电话：${s.phone || ''}</span>
                        <span>日期：${s.date || ''}</span>
                    </div>
                    <table class="sales-print-table">
                        <thead>
                            <tr>
                                <th style="width:12%;">序号</th>
                                <th style="width:22%;">货品名称</th>
                                <th style="width:11%;">规格</th>
                                <th style="width:12%;">数量</th>
                                <th style="width:9%;">单位</th>
                                <th style="width:13%;">单价</th>
                                <th style="width:13%;">金额</th>
                                <th style="width:8%;">备注</th>
                            </tr>
                        </thead>
                        <tbody>${rowsHtml}</tbody>
                    </table>
                    <div class="sale-total-row">
                        <span>合计人民币：${Utils.numberToChinese(s.total)}</span>
                        <span>合计金额：¥${Number(s.total || 0).toFixed(2)}</span>
                    </div>
                    <div class="sale-sign-row">
                        <span>客户签字：__________</span>
                        <span>单号：${s.no || ''}</span>
                    </div>
                    ${pageIndex === pages.length - 1 ? handlerRow : ''}
                </section>
            `;
        }).join('');
    },

    async showSales(id, autoPrint = false) {
        const s = await Sales.ensureDetail(id);
        if (!s) return;
        this.openDocument(this.renderSalesDocument(s), autoPrint, 'TRIPLICATE_HALF');
    },

    async printSalesBatch(ids) {
        const sales = (await Promise.all(ids.map(id => Sales.ensureDetail(id))))
            .filter(Boolean)
            .sort((a, b) => {
                const dateCompare = String(a.date || '').localeCompare(String(b.date || ''));
                if (dateCompare !== 0) return dateCompare;
                return String(a.no || '').localeCompare(String(b.no || ''));
            });
        let pageIndex = 0;
        const html = sales.map(sale => {
            const pageCount = Math.max(1, Math.ceil((sale.items || []).length / 20));
            const saleHtml = this.renderSalesDocument(sale, pageIndex);
            pageIndex += pageCount;
            return saleHtml;
        }).join('');
        this.openDocument(html, true, 'TRIPLICATE_HALF');
    },

    async showSalesDate(date, autoPrint = false) {
        const sales = await Promise.all(Store.sales
            .filter(s => (s.date || '未填写日期') === date)
            .map(sale => Sales.ensureDetail(sale.id)));
        if (sales.length === 0) return;

        const salesHtml = sales
            .sort((a, b) => String(a.no || '').localeCompare(String(b.no || '')))
            .map((sale, index) => this.renderSalesDocument(sale, index))
            .join('');

        this.openDocument(salesHtml, autoPrint, 'TRIPLICATE_HALF');
    },

    doPrint(html) {
        this.openDocument(html, true);
    },

    openDocument(html, autoPrint = false, pageSize = 'A4') {
        const isTriplicateHalfPage = pageSize === 'TRIPLICATE_HALF';
        const isPurchaseA4 = pageSize === 'PURCHASE_A4';
        const settings = this.getPrintSettings();
        const salesContentWidth = Math.max(180, settings.salesWidth - settings.salesMarginLeft - settings.salesMarginRight);
        const salesContentHeight = Math.max(100, settings.salesHeight - settings.salesMarginTop - settings.salesMarginBottom);
        const purchaseContentWidth = Math.max(160, 210 - settings.purchaseMargin * 2);
        const printCss = isTriplicateHalfPage
            ? `
                    @page {
                        size: ${settings.salesWidth}mm ${settings.salesHeight}mm;
                        margin: ${settings.salesMarginTop}mm ${settings.salesMarginRight}mm ${settings.salesMarginBottom}mm ${settings.salesMarginLeft}mm;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                        font-family: "Microsoft YaHei", "SimSun", sans-serif;
                        color: #000;
                    }
                    .sales-print-page {
                        width: ${salesContentWidth}mm;
                        min-height: ${salesContentHeight}mm;
                        box-sizing: border-box;
                        padding: 0;
                        overflow: hidden;
                    }
                    .page-break { page-break-before: always; break-before: page; }
                    .sales-print-page h2 {
                        margin: 0 0 2mm;
                        text-align: center;
                        font-size: 15pt;
                        line-height: 1.15;
                        font-weight: 700;
                    }
                    .sale-meta {
                        font-size: 9.5pt;
                        line-height: 1.2;
                        margin-bottom: 1mm;
                    }
                    .sale-meta-grid {
                        display: grid;
                        grid-template-columns: 1.6fr .8fr .8fr;
                        gap: 2mm;
                    }
                    .sales-print-table {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                        font-size: ${settings.salesFontSize}pt;
                    }
                    .sales-print-table th,
                    .sales-print-table td {
                        border: 1px solid #000;
                        padding: 0.45mm 0.8mm;
                        height: ${settings.salesRowHeight}mm;
                        line-height: 1.05;
                        overflow: hidden;
                        word-break: break-all;
                    }
                    .sales-print-table th {
                        font-weight: 700;
                        text-align: center;
                    }
                    .sales-print-table td.num {
                        text-align: right;
                        word-break: normal;
                    }
                    .sale-total-row {
                        display: grid;
                        grid-template-columns: 1fr 54mm;
                        border: 1px solid #000;
                        border-top: 0;
                        font-size: 9.5pt;
                        line-height: 6mm;
                        min-height: 6mm;
                    }
                    .sale-total-row span {
                        padding: 0 2mm;
                    }
                    .sale-total-row span + span {
                        border-left: 1px solid #000;
                    }
                    .sale-sign-row {
                        margin-top: 2mm;
                        display: flex;
                        justify-content: space-between;
                        gap: 10mm;
                        font-size: 9.5pt;
                        line-height: 1.2;
                    }
                    @media print {
                        table { page-break-inside: auto; }
                        tr { page-break-inside: avoid; page-break-after: auto; }
                    }
                `
            : isPurchaseA4
            ? `
                    @page { size: A4 portrait; margin: ${settings.purchaseMargin}mm; }
                    html, body {
                        margin: 0;
                        padding: 0;
                        font-family: "Microsoft YaHei", "SimSun", sans-serif;
                        color: #000;
                    }
                    .purchase-print-page {
                        box-sizing: border-box;
                        width: ${purchaseContentWidth}mm;
                    }
                    .purchase-print-page h2 {
                        margin: 0 0 6mm;
                        text-align: center;
                        font-size: 18pt;
                        line-height: 1.2;
                        font-weight: 700;
                    }
                    .purchase-meta {
                        display: flex;
                        justify-content: space-between;
                        margin-bottom: 4mm;
                        font-size: 11pt;
                    }
                    .purchase-print-table {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                        font-size: ${settings.purchaseFontSize}pt;
                    }
                    .purchase-print-table th,
                    .purchase-print-table td {
                        border: 1px solid #000;
                        padding: 1.8mm 2mm;
                        line-height: 1.25;
                        word-break: break-all;
                    }
                    .purchase-print-table th {
                        text-align: center;
                        font-weight: 700;
                    }
                    .purchase-print-table td.num {
                        text-align: right;
                        word-break: normal;
                    }
                    .purchase-sign-row {
                        margin-top: 10mm;
                        display: flex;
                        justify-content: space-between;
                        gap: 10mm;
                        font-size: 11pt;
                    }
                    @media print {
                        table { page-break-inside: auto; }
                        tr { page-break-inside: avoid; page-break-after: auto; }
                        thead { display: table-header-group; }
                    }
                `
            : `
                    @media print {
                        body { margin: 0; padding: 10mm; }
                        table { page-break-inside: auto; }
                        tr { page-break-inside: avoid; page-break-after: auto; }
                    }
                `;
        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <html>
            <head>
                <title>打印</title>
                <style>
                    ${printCss}
                </style>
            </head>
            <body>${html}</body>
            </html>
        `);
        printWindow.document.close();
        printWindow.focus();
        if (autoPrint) {
            setTimeout(() => printWindow.print(), 200);
        }
    }
};

// ==================== Excel导入导出 ====================
const ExcelImport = {
    openHotelSalesImport() {
        document.getElementById('hotelSalesTextDate').value = Utils.today();
        document.getElementById('hotelGrocerySupplier').value = '西安禾润佳商贸有限公司';
        this.renderHotelCustomerOptions();
        const defaultCustomer = Store.customers.find(c => c.name === '西安汉庭酒店（大明宫万达）') || Store.customers[0];
        document.getElementById('hotelSalesCustomerSelect').value = defaultCustomer ? defaultCustomer.id : '';
        document.getElementById('hotelSalesTextCustomer').value = defaultCustomer ? defaultCustomer.name : '西安汉庭酒店（大明宫万达）';
        document.getElementById('hotelSalesImportResult').style.display = 'none';
        Modals.open('hotelSalesImportModal');
    },

    renderHotelCustomerOptions() {
        const select = document.getElementById('hotelSalesCustomerSelect');
        if (!select) return;
        select.innerHTML = '<option value="">手动输入客户</option>' +
            Store.customers.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
    },

    onHotelCustomerChange() {
        const customerId = document.getElementById('hotelSalesCustomerSelect').value;
        if (!customerId) return;
        const customer = Store.find('customers', customerId);
        if (customer) {
            document.getElementById('hotelSalesTextCustomer').value = customer.name;
        }
    },

    downloadProductTemplate() {
        const csv = '﻿编号,名称,规格,单位,采购单价,初始库存,预警线,专属部门\nA001,打印纸A4,70g 500张/包,包,25.00,0,20,办公\nA002,大米,50斤/袋,袋,120.00,0,10,食堂\n';
        Utils.downloadFile(csv, '物品导入模板.csv');
    },

    downloadOutboundTemplate() {
        const csv = '﻿日期,部门,领用人,物品编号,物品名称,数量,单价\n2026-05-14,食堂,张三,A001,打印纸A4,10,25.00\n2026-05-14,食堂,张三,A002,大米,5,120.00\n';
        Utils.downloadFile(csv, '领用单导入模板.csv');
    },

    downloadPurchaseTemplate() {
        const csv = '﻿日期,供应商,物品编号,物品名称,规格,单位,数量,单价\n2026-05-14,XX供应商,A001,打印纸A4,70g 500张/包,包,100,25.00\n2026-05-14,XX供应商,A002,大米,50斤/袋,袋,50,120.00\n';
        Utils.downloadFile(csv, '采购单导入模板.csv');
    },

    downloadHotelSalesTemplate() {
        const csv = '﻿日期,酒店名称,物品编号,物品名称,规格,单位,数量\n2026-05-22,酒店,A001,打印纸A4,70g 500张/包,包,10\n2026-05-22,酒店,A002,大米,50斤/袋,袋,5\n';
        Utils.downloadFile(csv, '酒店清单生成销售单模板.csv');
    },

    async importProducts(input) {
        const file = input.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const result = await API.upload('import/products-excel', formData);
            const resultDiv = document.getElementById('excelImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `✅ 导入成功 ${result.imported} 条，跳过 ${result.skipped} 条（编号重复）`;
            await Store.refresh('products');
            Products.render();
        } catch (e) {
            const resultDiv = document.getElementById('excelImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `❌ 导入失败：${e.message}`;
        }
        input.value = '';
    },

    async importSalesPrices(input) {
        const file = input.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);
        const fallbackCustomer = document.getElementById('salesPriceImportFallbackCustomer')?.value?.trim();
        if (fallbackCustomer) formData.append('customer', fallbackCustomer);

        const resultDiv = document.getElementById('salesPriceImportResult');
        try {
            const result = await API.upload('import/sales-prices-excel', formData);
            const samples = result.samples || [];
            const customers = result.customers || [];
            const sampleHtml = samples.length
                ? `<div style="margin-top:8px;color:#666;">示例：${samples.map(s => `${s.customer}：${s.name}/${s.unit} ¥${Number(s.price || 0).toFixed(2)}`).join('；')}</div>`
                : '';
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = [
                `导入成功：读取 ${result.sourceRows || 0} 条销售明细，整理出 ${result.latestPrices || 0} 条客户售价`,
                `新增售价 ${result.createdPrices || 0} 条，更新售价 ${result.updatedPrices || 0} 条`,
                result.createdProducts ? `自动新增物品 ${result.createdProducts} 个` : '',
                customers.length ? `客户：${customers.join('、')}` : '',
                result.skipped ? `跳过 ${result.skipped} 条无单价明细` : ''
            ].filter(Boolean).join('<br>') + sampleHtml;
            await Store.refresh(['products', 'customers']);
            Products.render();
        } catch (e) {
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `导入失败：${e.message}`;
        }
        input.value = '';
    },

    async processUploadedSalesPrices() {
        const resultDiv = document.getElementById('uploadedExcelProcessResult');
        if (resultDiv) {
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f5f5f5';
            resultDiv.innerHTML = '正在处理 WinSCP 上传的 Excel...';
        }
        try {
            const result = await API.post('uploaded-excels/process', {});
            const failed = (result.files || []).filter(file => !file.ok);
            const success = (result.files || []).filter(file => file.ok);
            const successHtml = success.length
                ? `<div style="margin-top:8px;color:#666;">成功文件：${success.map(file => file.name).join('、')}</div>`
                : '';
            const failedHtml = failed.length
                ? `<div style="margin-top:8px;color:#a8071a;">失败文件：${failed.map(file => `${file.name}（${file.error || '未知错误'}）`).join('；')}</div>`
                : '';
            if (resultDiv) {
                resultDiv.style.background = result.failed ? '#fff1f0' : '#f6ffed';
                resultDiv.innerHTML = [
                    `扫描 ${result.scanned || 0} 个文件，成功 ${result.processed || 0} 个，失败 ${result.failed || 0} 个`,
                    `读取 ${result.sourceRows || 0} 条销售明细，整理出 ${result.latestPrices || 0} 条客户售价`,
                    `新增售价 ${result.createdPrices || 0} 条，更新售价 ${result.updatedPrices || 0} 条`,
                    result.createdProducts ? `自动新增物品 ${result.createdProducts} 个` : '',
                    (result.customers || []).length ? `客户：${result.customers.join('、')}` : ''
                ].filter(Boolean).join('<br>') + successHtml + failedHtml;
            }
            await Store.refresh(['products', 'customers']);
            Products.render();
        } catch (e) {
            if (resultDiv) {
                resultDiv.style.background = '#fff1f0';
                resultDiv.innerHTML = `处理失败：${e.message}`;
            } else {
                alert(`处理失败：${e.message}`);
            }
        }
    },

    async importOutbounds(input) {
        const file = input.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const result = await API.upload('import/outbounds-excel', formData);
            const resultDiv = document.getElementById('outboundImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `✅ 成功生成 ${result.created} 张领用单`;
            await Store.refresh(['products', 'outbounds']);
            Outbounds.render();
            Dashboard.render();
        } catch (e) {
            const resultDiv = document.getElementById('outboundImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `❌ 导入失败：${e.message}`;
        }
        input.value = '';
    },

    async importPurchases(input) {
        const file = input.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const result = await API.upload('import/purchases-excel', formData);
            const resultDiv = document.getElementById('purchaseImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `✅ 导入成功`;
            await Store.refresh(['products', 'purchases']);
            Purchases.render();
            Dashboard.render();
        } catch (e) {
            const resultDiv = document.getElementById('purchaseImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `❌ 导入失败：${e.message}`;
        }
        input.value = '';
    },

    async importHotelSales(input) {
        const file = input.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const result = await API.upload('import/hotel-sales-excel', formData);
            const resultDiv = document.getElementById('hotelSalesImportResult');
            const noPrice = result.noLastPrice || [];
            const missing = result.missingProducts || [];
            const extra = [
                noPrice.length ? `有 ${noPrice.length} 个物品没有上次销售价，已临时使用采购单价` : '',
                missing.length ? `未匹配物品：${missing.slice(0, 8).join('、')}${missing.length > 8 ? ' 等' : ''}` : ''
            ].filter(Boolean).join('<br>');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `生成成功：${result.created} 张销售单，${result.items} 条明细，扣库存 ${result.deductedItems || 0} 条，跳过 ${result.skipped} 条${extra ? '<br>' + extra : ''}`;
            await Store.refresh(['products', 'sales']);
            Sales.render();
            Dashboard.render();
        } catch (e) {
            const resultDiv = document.getElementById('hotelSalesImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `导入失败：${e.message}`;
        }
        input.value = '';
    },

    async importHotelSalesFile(input) {
        const file = input.files[0];
        if (!file) return;

        if (file.name.toLowerCase().endsWith('.xlsx')) {
            await this.importHotelOrdersXlsx(input);
        } else {
            await this.importHotelSales(input);
        }
    },

    async importHotelOrdersXlsx(input) {
        const file = input.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const result = await API.upload('import/hotel-orders-xlsx', formData);
            const resultDiv = document.getElementById('hotelSalesImportResult');
            const noPrice = result.noPrice || [];
            const extra = [
                result.createdProducts ? `自动新增 ${result.createdProducts} 个酒店流转物品` : '',
                result.duplicateOrders ? `跳过 ${result.duplicateOrders} 张重复订单` : '',
                noPrice.length ? `无价格物品：${noPrice.slice(0, 8).join('、')}${noPrice.length > 8 ? ' 等' : ''}` : ''
            ].filter(Boolean).join('<br>');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `导入成功：${result.orders} 个订单块，生成 ${result.created} 张销售单，${result.items} 条明细，扣库存 ${result.deductedItems || 0} 条${extra ? '<br>' + extra : ''}`;
            await Store.refresh(['products', 'sales']);
            Sales.render();
            Dashboard.render();
        } catch (e) {
            const resultDiv = document.getElementById('hotelSalesImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `导入失败：${e.message}`;
        }
        input.value = '';
    },

    async importHotelSalesText() {
        const text = document.getElementById('hotelSalesText').value.trim();
        if (!text) {
            Utils.toast('请先粘贴酒店清单', 'error');
            return;
        }

        try {
            const result = await API.post('import/hotel-sales-text', {
                text,
                date: document.getElementById('hotelSalesTextDate').value,
                customer: document.getElementById('hotelSalesTextCustomer').value.trim() || '酒店'
            });
            const resultDiv = document.getElementById('hotelSalesImportResult');
            const noPrice = result.noLastPrice || [];
            const missing = result.missingProducts || [];
            const unparsed = result.unparsed || [];
            const extra = [
                noPrice.length ? `有 ${noPrice.length} 个物品没有上次销售价，已临时使用采购单价` : '',
                missing.length ? `未匹配物品：${missing.slice(0, 8).join('、')}${missing.length > 8 ? ' 等' : ''}` : '',
                unparsed.length ? `未识别片段：${unparsed.slice(0, 8).join('、')}${unparsed.length > 8 ? ' 等' : ''}` : ''
            ].filter(Boolean).join('<br>');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `生成成功：${result.created} 张销售单，${result.items} 条明细，扣库存 ${result.deductedItems || 0} 条，跳过 ${result.skipped} 条${extra ? '<br>' + extra : ''}`;
            await Store.refresh(['products', 'sales']);
            Sales.render();
            Dashboard.render();
        } catch (e) {
            const resultDiv = document.getElementById('hotelSalesImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `生成失败：${e.message}`;
        }
    },

    async syncHotelGroceryText() {
        const text = document.getElementById('hotelSalesText').value.trim();
        if (!text) {
            Utils.toast('请先粘贴酒店订单', 'error');
            return;
        }

        try {
            const result = await API.post('hotel-grocery/sync', {
                text,
                date: document.getElementById('hotelSalesTextDate').value,
                supplierName: document.getElementById('hotelGrocerySupplier').value.trim() || '西安禾润佳商贸有限公司',
                customer: document.getElementById('hotelSalesTextCustomer').value.trim() || '西安汉庭酒店（大明宫万达）'
            });
            const resultDiv = document.getElementById('hotelSalesImportResult');
            const unparsed = result.unparsed || [];
            const created = result.createdProductNames || [];
            const extra = [
                result.saleNo ? `销售单：${result.saleNo}` : '',
                result.smartParsed ? `方舟智能识别：${result.smartParsed} 条` : '',
                created.length ? `自动新增物品：${created.slice(0, 8).join('、')}${created.length > 8 ? ' 等' : ''}` : '',
                unparsed.length ? `未识别片段：${unparsed.slice(0, 8).join('、')}${unparsed.length > 8 ? ' 等' : ''}` : ''
            ].filter(Boolean).join('<br>');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = `生成成功：生成 ${result.created || 1} 张销售单，${result.items} 条明细，扣库存 ${result.deductedItems || 0} 条${extra ? '<br>' + extra : ''}`;
            await Store.refresh(['products', 'sales']);
            Sales.render();
            Dashboard.render();
        } catch (e) {
            const resultDiv = document.getElementById('hotelSalesImportResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = `同步失败：${e.message}`;
        }
    },

    async downloadHotelGroceryPurchaseList() {
        const text = document.getElementById('hotelSalesText').value.trim();
        if (!text) {
            Utils.toast('请先粘贴酒店订单', 'error');
            return;
        }

        try {
            const date = document.getElementById('hotelSalesTextDate').value || Utils.today();
            const customerSelect = document.getElementById('hotelSalesCustomerSelect');
            const customerName = customerSelect && customerSelect.value
                ? (customerSelect.options[customerSelect.selectedIndex]?.textContent || '').trim()
                : (document.getElementById('hotelSalesTextCustomer').value || '酒店').trim();
            const res = await fetch(`${API_BASE}/api/hotel-grocery/purchase-list`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    text,
                    date,
                    title: `${customerName || '酒店'}采购填表单${date}`
                })
            });
            if (!res.ok) {
                const result = await res.json();
                throw new Error(result.error || `HTTP ${res.status}`);
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${customerName || '酒店'}采购单_${date}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
            Utils.toast('采购单已生成');
        } catch (e) {
            Utils.toast('生成采购单失败：' + e.message, 'error');
        }
    }
};

// ==================== 数据备份 ====================
const Backup = {
    async render() {
        this.checkAutoBackup();
    },

    checkAutoBackup() {
        const now = new Date();
        const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const lastBackup = localStorage.getItem('inv_last_backup_month');

        const statusDiv = document.getElementById('autoBackupStatus');
        if (lastBackup === currentMonth) {
            statusDiv.innerHTML = '✅ ' + currentMonth + ' 已备份';
            statusDiv.style.background = '#f6ffed';
            statusDiv.style.color = '#52c41a';
        } else {
            statusDiv.innerHTML = '⚠️ 本月尚未备份';
            statusDiv.style.background = '#fffbe6';
            statusDiv.style.color = '#fa8c16';
            // 自动备份
            this.manualBackup();
        }
    },

    async manualBackup() {
        try {
            const data = await API.get('backup');
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const now = new Date();
            const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
            const a = document.createElement('a');
            a.href = url;
            a.download = `inventory_auto_backup_${month}.json`;
            a.click();
            URL.revokeObjectURL(url);

            localStorage.setItem('inv_last_backup_month', month);
            this.checkAutoBackup();
            Utils.toast('备份成功');
        } catch (e) {
            Utils.toast('备份失败：' + e.message, 'error');
        }
    },

    async exportData() {
        try {
            const data = await API.get('backup');
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `inventory_backup_${Utils.today()}.json`;
            a.click();
            URL.revokeObjectURL(url);
            Utils.toast('导出成功');
        } catch (e) {
            Utils.toast('导出失败：' + e.message, 'error');
        }
    },

    async importData(input) {
        const file = input.files[0];
        if (!file) return;

        if (!confirm('导入将覆盖现有数据，确定继续？')) {
            input.value = '';
            return;
        }

        try {
            const text = await file.text();
            const data = JSON.parse(text);
            await API.post('backup', data);
            await Store.loadAll();

            const resultDiv = document.getElementById('importResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#f6ffed';
            resultDiv.innerHTML = '✅ 数据导入成功';

            // 刷新当前页面
            App.navigate(App.currentPage);
            Utils.toast('导入成功');
        } catch (e) {
            const resultDiv = document.getElementById('importResult');
            resultDiv.style.display = 'block';
            resultDiv.style.background = '#fff1f0';
            resultDiv.innerHTML = '❌ 导入失败：' + e.message;
        }
        input.value = '';
    },

    async clearAll() {
        if (!confirm('⚠️ 确定清空所有数据？此操作不可恢复！')) return;
        if (!confirm('再次确认：将删除所有物品、采购单、领用单、销售单、财务记录！')) return;

        try {
            await API.post('clear', {});
            await Store.loadAll();
            App.navigate('dashboard');
            Utils.toast('数据已清空');
        } catch (e) {
            Utils.toast('清空失败：' + e.message, 'error');
        }
    }
};

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});
