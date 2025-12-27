// Mobile JS for BestOne DB

// Data & State
let enhancedData = [];
let currentState = {
    search: '',
    filterLevel: 'all',
    filterSubject: 'all',
    filterSpecial: 'all', // 'all' only for now in mobile UI simplification
    tab: 'home', // home, favorites, cart

    // Core Data
    students: [],
    currentStudentId: null,
    favorites: [],
    history: [],
    settings: {
        schoolName: 'ECCベストワン藍住・北島中央',
        taxRate: 0.10
    }
};

// Utils: Tag Estimation (Copied from script.js)
function estimateTagsFromTitle(title) {
    if (!title) return { level: 'unknown', subject: 'unknown', special: [] };
    let level = 'unknown', subject = 'unknown', special = [];
    const t = title.toLowerCase();

    if (t.includes('小') || t.includes('小学')) level = 'elementary';
    if (t.includes('中') || t.includes('中学') || t.includes('高校入試')) level = 'junior';
    if (t.includes('高') || t.includes('高校') || t.includes('大学入試')) level = 'high';

    if (t.includes('英')) subject = 'english';
    if (t.includes('数') || t.includes('算')) subject = 'math';
    if (t.includes('国') || t.includes('文') || t.includes('漢字')) subject = 'japanese';
    if (t.includes('理') || t.includes('物') || t.includes('化') || t.includes('生')) subject = 'science';
    if (t.includes('社') || t.includes('史') || t.includes('地') || t.includes('公')) subject = 'social';

    if (t.includes('英検')) special.push('eiken');
    if (t.includes('受験') || t.includes('入試')) special.push('exam');
    if (t.includes('講習')) special.push('season');

    return { level, subject, special };
}

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    // 1. Data Init
    const dataSource = (typeof DB_DATA !== 'undefined') ? DB_DATA : [];
    enhancedData = dataSource.map(item => ({ ...item, ...estimateTagsFromTitle(item.title) }));

    // 2. Load State
    loadState();

    // 3. Render Initial
    renderHeaderStudent();
    renderContent();
    updateCartBadge();

    // 4. Events
    setupEventListeners();
});

// State Management
function loadState() {
    const saved = localStorage.getItem('materio_db_state');
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            currentState.students = parsed.students || [];
            currentState.currentStudentId = parsed.currentStudentId || null;
            currentState.favorites = parsed.favorites || [];
            currentState.history = parsed.history || [];
            if (parsed.settings) currentState.settings = parsed.settings;
        } catch (e) { console.error('Load Error', e); }
    }
}

function saveState() {
    localStorage.setItem('materio_db_state', JSON.stringify({
        students: currentState.students,
        currentStudentId: currentState.currentStudentId,
        favorites: currentState.favorites,
        history: currentState.history,
        settings: currentState.settings
    }));
}

// Rendering

function renderContent() {
    const container = document.getElementById('itemList');
    container.innerHTML = '';

    // Determine what to show based on Tab
    if (currentState.tab === 'home') {
        renderHomeItems(container);
    } else if (currentState.tab === 'favorites') {
        renderFavorites(container);
    } else if (currentState.tab === 'cart') {
        renderCart(container);
    }
}

function renderHomeItems(container) {
    // Filters
    let items = enhancedData.filter(item => {
        // Search
        if (currentState.search) {
            const keywords = currentState.search.toLowerCase().split(/\s+/).filter(k => k);
            if (!keywords.every(k => item.title.toLowerCase().includes(k))) return false;
        }
        // Category Filters
        if (currentState.filterLevel !== 'all' && item.level !== currentState.filterLevel) return false;
        if (currentState.filterSubject !== 'all' && item.subject !== currentState.filterSubject) return false;

        return true;
    });

    // Limit to 50 for performance if searching/all
    const displayItems = items.slice(0, 50);

    if (displayItems.length === 0) {
        container.innerHTML = '<div class="empty-state"><i class="fa-solid fa-magnifying-glass"></i><p>該当する教材がありません</p></div>';
        return;
    }

    // Render List
    displayItems.forEach(item => {
        container.appendChild(createItemCard(item));
    });

    if (items.length > 50) {
        const more = document.createElement('div');
        more.style.textAlign = 'center';
        more.style.padding = '10px';
        more.style.color = '#999';
        more.innerText = `他 ${items.length - 50} 件 (検索して絞り込んでください)`;
        container.appendChild(more);
    }
}

function renderFavorites(container) {
    const favItems = enhancedData.filter(item => currentState.favorites.includes(item.id));

    if (favItems.length === 0) {
        container.innerHTML = '<div class="empty-state"><i class="fa-regular fa-star"></i><p>お気に入りはまだありません</p></div>';
        return;
    }

    favItems.forEach(item => {
        container.appendChild(createItemCard(item));
    });
}

function renderCart(container) {
    // Check student
    if (!currentState.currentStudentId) {
        container.innerHTML = `
            <div class="empty-state">
                <i class="fa-solid fa-user-xmark"></i>
                <p>生徒が選択されていません</p>
                <button class="btn-sm" style="margin-top:10px; font-size:16px;" onclick="openStudentModal()">生徒を選択</button>
            </div>`;
        // Hide Footer
        removeCartFooter();
        return;
    }

    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (!student || student.cart.length === 0) {
        container.innerHTML = '<div class="empty-state"><i class="fa-solid fa-cart-shopping"></i><p>カートは空です</p></div>';
        removeCartFooter();
        return;
    }

    // Cart Items
    // Add extra padding at bottom for the fixed footer
    container.style.paddingBottom = '100px';

    student.cart.forEach(item => {
        const div = document.createElement('div');
        div.className = 'item-card';
        div.innerHTML = `
            <div class="item-main">
                <div class="item-title">${item.title}</div>
                <div class="item-price">¥${item.price_retail.toLocaleString()}</div>
            </div>
            <div class="item-actions">
                <button class="btn-sm" style="background:#fee2e2; color:#ef4444;" onclick="removeFromCart('${item.id}')">
                    <i class="fa-solid fa-trash"></i> 削除
                </button>
            </div>
        `;
        container.appendChild(div);
    });

    // Render Fixed Footer
    renderCartFooter(student);
}

function removeCartFooter() {
    const existing = document.querySelector('.cart-footer-actions');
    if (existing) existing.remove();
}

function renderCartFooter(student) {
    removeCartFooter();
    const total = student.cart.reduce((sum, i) => sum + (parseInt(i.price_retail) || 0), 0);

    const footer = document.createElement('div');
    footer.className = 'cart-footer-actions';
    footer.innerHTML = `
        <div class="cart-total">
            <span>合計</span>
            <span>¥${total.toLocaleString()}</span>
        </div>
        <button class="btn-checkout" id="btnCheckout">見積書を作成</button>
    `;

    document.querySelector('.app-container').appendChild(footer);

    document.getElementById('btnCheckout').addEventListener('click', () => {
        // Same print logic as desktop (simplified)
        printQuotation(student);
    });
}

function createItemCard(item) {
    const div = document.createElement('div');
    div.className = 'item-card';
    div.onclick = () => openDetailModal(item);

    const isFav = currentState.favorites.includes(item.id);
    const favIcon = isFav ? '<i class="fa-solid fa-star" style="color:#fbbf24"></i>' : '';

    div.innerHTML = `
        <div class="item-main">
            <div class="item-title">${item.title}</div>
            <div class="item-meta">
                <span class="badge level-${item.level}">${getLevelLabel(item.level)}</span>
                <span class="badge subject-${item.subject}">${getSubjectLabel(item.subject)}</span>
                ${favIcon}
            </div>
            <div class="item-price">¥${item.price_retail.toLocaleString()}</div>
        </div>
        <div class="item-actions">
            <button class="btn-add" onclick="event.stopPropagation(); quickAdd('${item.id}')">
                <i class="fa-solid fa-plus"></i>
            </button>
        </div>
    `;
    return div;
}

// Business Logic
function quickAdd(itemId) {
    const item = enhancedData.find(i => i.id === itemId);
    if (!item) return;

    if (!currentState.currentStudentId) {
        alert('生徒を選択してください');
        openStudentModal();
        return;
    }

    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (student.cart.some(c => c.id === item.id)) {
        alert('既にカートに入っています');
        return;
    }

    student.cart.push(item);
    saveState();
    updateCartBadge();

    // Animation/Feedback
    alert(`${item.title}\nをカートに追加しました`);
}

function removeFromCart(itemId) {
    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (student) {
        student.cart = student.cart.filter(i => i.id !== itemId);
        saveState();
        updateCartBadge();
        renderContent(); // Re-render cart
    }
}

function updateCartBadge() {
    const badge = document.getElementById('cartBadge');
    if (!currentState.currentStudentId) {
        badge.style.display = 'none';
        return;
    }
    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    const count = student ? student.cart.length : 0;

    if (count > 0) {
        badge.textContent = count;
        badge.style.display = 'flex';
    } else {
        badge.style.display = 'none';
    }
}

function renderHeaderStudent() {
    const banner = document.getElementById('studentBanner');
    const nameData = document.getElementById('currentStudentName');

    if (!currentState.currentStudentId) {
        nameData.textContent = '選択されていません';
        nameData.style.color = '#999';
    } else {
        const student = currentState.students.find(s => s.id === currentState.currentStudentId);
        if (student) {
            nameData.textContent = student.name + (student.grade ? ` (${student.grade})` : '');
            nameData.style.color = '#333';
        }
    }
}

// View Controllers
function openStudentModal() {
    const list = document.getElementById('studentList');
    list.innerHTML = '';

    currentState.students.forEach(student => {
        const div = document.createElement('div');
        div.className = `student-item ${student.id === currentState.currentStudentId ? 'active' : ''}`;
        div.innerHTML = `
            <span>${student.name} ${student.grade ? `(${student.grade})` : ''}</span>
            ${student.id === currentState.currentStudentId ? '<i class="fa-solid fa-check"></i>' : ''}
        `;
        div.onclick = () => {
            currentState.currentStudentId = student.id;
            saveState();
            renderHeaderStudent();
            updateCartBadge();
            closeStudentModal();
            // If we were in cart, re-render
            if (currentState.tab === 'cart') renderContent();
        };
        list.appendChild(div);
    });

    document.getElementById('studentModal').classList.add('active');
}

function closeStudentModal() {
    document.getElementById('studentModal').classList.remove('active');
}

function openDetailModal(item) {
    document.getElementById('detailTitle').textContent = item.title;
    document.getElementById('detailRetail').textContent = `¥${item.price_retail.toLocaleString()}`;
    document.getElementById('detailWholesale').textContent = `¥${item.price_wholesale.toLocaleString()}`;

    // Badges
    const badgeContainer = document.getElementById('detailBadges');
    badgeContainer.innerHTML = '';
    // ... add badges ...

    // Fav Button
    const favBtn = document.getElementById('detailFavBtn');
    const isFav = currentState.favorites.includes(item.id);
    favBtn.innerHTML = isFav ? '<i class="fa-solid fa-star"></i>' : '<i class="fa-regular fa-star"></i>';
    favBtn.onclick = () => {
        if (isFav) {
            currentState.favorites = currentState.favorites.filter(id => id !== item.id);
        } else {
            currentState.favorites.push(item.id);
        }
        saveState();
        openDetailModal(item); // Update View
        if (currentState.tab === 'favorites') renderContent();
    };

    // Add Cart Button
    const addBtn = document.getElementById('detailAddCartBtn');
    addBtn.onclick = () => {
        quickAdd(item.id);
        document.getElementById('detailModal').classList.remove('active');
    };

    document.getElementById('detailModal').classList.add('active');
}

// Event Setup
function setupEventListeners() {
    // 1. Bottom Nav
    document.querySelectorAll('.nav-item').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const tab = btn.dataset.tab;
            if (tab === 'menu') {
                document.getElementById('menuModal').classList.add('active');
            } else {
                // Switch Tab
                document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                currentState.tab = tab;
                renderContent();
            }
        });
    });

    // 2. Search
    const searchInput = document.getElementById('searchInput');
    const clearBtn = document.getElementById('clearSearchBtn');

    searchInput.addEventListener('input', (e) => {
        currentState.search = e.target.value;
        currentState.tab = 'home'; // Force switch to home
        document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
        document.querySelector('[data-tab="home"]').classList.add('active');

        clearBtn.style.display = currentState.search ? 'block' : 'none';
        renderContent();
    });

    clearBtn.addEventListener('click', () => {
        currentState.search = '';
        searchInput.value = '';
        clearBtn.style.display = 'none';
        renderContent();
    });

    // 3. Filters
    document.querySelectorAll('.chip').forEach(chip => {
        chip.addEventListener('click', () => {
            const filterType = chip.dataset.filter;

            // Check which category this filter belongs to
            // Simple logic for now: predefined lists
            const levels = ['all', 'elementary', 'junior', 'high'];
            const subjects = ['english', 'math', 'japanese', 'science', 'social'];

            if (levels.includes(filterType)) {
                // Reset other levels
                levels.forEach(l => {
                    document.querySelector(`[data-filter="${l}"]`)?.classList.remove('active');
                });
                currentState.filterLevel = filterType;
            } else if (subjects.includes(filterType)) {
                // Reset other subjects (need 'all' for subject too in a robust app, but UI simplified)
                subjects.forEach(s => {
                    document.querySelector(`[data-filter="${s}"]`)?.classList.remove('active');
                });
                currentState.filterSubject = filterType;
            } else if (filterType === 'all') {
                // Global All basically
                currentState.filterLevel = 'all';
                currentState.filterSubject = 'all';
                document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
                // Re-add active to relevant ones
            }

            chip.classList.add('active');

            // Handle cross-group "All" logic better:
            if (filterType === 'all') {
                // Reset visuals
                document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
                chip.classList.add('active');
            }

            currentState.tab = 'home';
            document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
            document.querySelector('[data-tab="home"]').classList.add('active');

            renderContent();
        });
    });

    // 4. Modals
    document.getElementById('changeStudentBtn').addEventListener('click', openStudentModal);

    document.querySelectorAll('.close-sheet').forEach(b => b.addEventListener('click', () => {
        document.querySelectorAll('.sheet-overlay').forEach(s => s.classList.remove('active'));
    }));

    document.querySelectorAll('.close-detail').forEach(b => b.addEventListener('click', () => {
        document.getElementById('detailModal').classList.remove('active');
    }));

    // Student Creation
    document.getElementById('createNewStudentBtn').addEventListener('click', () => {
        document.getElementById('studentModal').classList.remove('active');
        document.getElementById('newStudentModal').classList.add('active');
    });

    document.getElementById('cancelNewStudentBtn').addEventListener('click', () => {
        document.getElementById('newStudentModal').classList.remove('active');
    });

    document.getElementById('saveNewStudentBtn').addEventListener('click', () => {
        const name = document.getElementById('newStudentNameInput').value;
        const grade = document.getElementById('newStudentGradeInput').value;
        if (name) {
            const newStudent = {
                id: Date.now().toString(),
                name,
                grade,
                cart: []
            };
            currentState.students.push(newStudent);
            currentState.currentStudentId = newStudent.id;
            saveState();
            renderHeaderStudent();
            renderContent(); // Updates if cart
            updateCartBadge();
            document.getElementById('newStudentModal').classList.remove('active');
        }
    });

    // Reset
    document.getElementById('menuReset').addEventListener('click', () => {
        if (confirm('全データを初期化しますか？')) {
            localStorage.removeItem('materio_db_state');
            location.reload();
        }
    });

    // Handle Settings
    document.getElementById('menuSettings').addEventListener('click', () => {
        const newName = prompt('教室名を入力してください:', currentState.settings.schoolName);
        if (newName !== null) {
            currentState.settings.schoolName = newName;
            saveState();
            alert('設定を保存しました');
        }
    });

    // Handle History
    document.getElementById('menuHistory').addEventListener('click', () => {
        let msg = '履歴(最新10件):\n';
        if (currentState.history.length === 0) msg = '履歴はありません';
        else {
            currentState.history.slice().reverse().slice(0, 10).forEach(h => {
                msg += `[${h.date}] ${h.type}: ${h.detail}\n`;
            });
        }
        alert(msg);
    });
}

// Helpers
function getLevelLabel(l) { return { elementary: '小', junior: '中', high: '高', unknown: '他' }[l] || l; }
function getSubjectLabel(s) { return { english: '英', math: '数', japanese: '国', science: '理', social: '社', unknown: '他' }[s] || s; }

function addToHistory(type, detail) {
    const now = new Date();
    const dateStr = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()} ${now.getHours()}:${now.getMinutes()}`;
    currentState.history.push({ date: dateStr, type, detail });
    if (currentState.history.length > 50) currentState.history.shift();
    saveState();
}

function printQuotation(student) {
    addToHistory('見積書作成', `${student.name}様の見積書を作成しました`);
    const today = new Date();
    const dateStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;

    let totalAmount = 0;
    const itemsHtml = student.cart.map((item, index) => {
        const price = parseInt(item.price_retail) || 0;
        totalAmount += price;
        return `
            <tr>
                <td class="center">${index + 1}</td>
                <td>${item.title}</td>
                <td class="center">1</td>
                <td class="right">¥${price.toLocaleString()}</td>
                <td class="right">¥${price.toLocaleString()}</td>
            </tr>
        `;
    }).join('');

    const ROW_TARGET = 10;
    const emptyRowsCount = Math.max(0, ROW_TARGET - student.cart.length);
    let emptyRowsHtml = '';
    for (let i = 0; i < emptyRowsCount; i++) {
        emptyRowsHtml += `
            <tr>
                <td class="center"></td>
                <td></td>
                <td class="center"></td>
                <td class="right"></td>
                <td class="right"></td>
            </tr>
        `;
    }

    const taxRate = currentState.settings.taxRate || 0.10;
    const taxAmount = Math.floor(totalAmount * taxRate / (1 + taxRate));

    const htmlContent = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <title>御見積書 - ${student.name}様</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;600&display=swap');
            body { font-family: 'Noto Serif JP', serif; margin: 0; padding: 0; background: #ccc; -webkit-print-color-adjust: exact; }
            .page { width: 210mm; height: 297mm; padding: 15mm; margin: 10mm auto; background: white; box-sizing: border-box; position: relative; }
            @media print { body { background: none; } .page { margin: 0; width: 100%; height: 100%; } @page { margin: 0; size: A4 portrait; } }
            .header { display: flex; justify-content: space-between; margin-bottom: 25px; }
            .title { font-size: 20pt; font-weight: 600; letter-spacing: 5px; border-bottom: 3px double #333; padding-bottom: 5px; }
            .date { text-align: right; font-size: 9pt; margin-bottom: 5px; }
            .info-block { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 25px; }
            .recipient { font-size: 14pt; border-bottom: 1px solid #333; padding-bottom: 3px; min-width: 250px; }
            .recipient span { font-size: 10pt; }
            .sender { font-size: 9pt; line-height: 1.4; text-align: right; }
            .company-name { font-size: 11pt; font-weight: 600; margin-bottom: 3px; }
            .total-block { margin-bottom: 20px; border-bottom: 2px solid #333; padding-bottom: 5px; }
            .total-label { font-size: 11pt; margin-right: 15px; }
            .total-amount { font-size: 18pt; font-weight: 600; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 10pt; }
            th { background-color: #f0f0f0; border: 1px solid #333; padding: 6px; font-weight: 600; text-align: center; }
            td { border: 1px solid #333; padding: 6px; height: 35px; }
            .center { text-align: center; }
            .right { text-align: right; }
            .col-no { width: 30px; } .col-qty { width: 40px; } .col-unit { width: 80px; } .col-amount { width: 80px; }
            .remarks { border: 1px solid #333; padding: 8px; height: 80px; font-size: 9pt; }
            .remarks-title { font-size: 9pt; font-weight: 600; margin-bottom: 3px; text-decoration: underline; }
        </style>
    </head>
    <body>
        <div class="page">
            <div class="date">${dateStr}</div>
            <div class="header"><div class="title">御見積書</div></div>
            <div class="info-block">
                <div class="recipient">${student.name} <span>様</span></div>
                <div class="sender">
                    <div class="company-name" style="font-size: 14pt;">${currentState.settings.schoolName}</div>
                    ${currentState.settings.address ? `<div>${currentState.settings.address}</div>` : ''}
                    ${currentState.settings.phone ? `<div>TEL: ${currentState.settings.phone}</div>` : ''}
                </div>
            </div>
            <div class="total-block">
                <span class="total-label">御見積金額</span>
                <span class="total-amount">¥${totalAmount.toLocaleString()}-</span>
                <span style="font-size: 10pt;"> (税込)</span>
                <div style="font-size: 9pt; text-align: right; margin-top: 5px; color: #555;">
                    (内消費税等 ${Math.round((currentState.settings.taxRate || 0.10) * 100)}%: ¥${taxAmount.toLocaleString()})
                </div>
            </div>
            <table>
                <thead>
                    <tr><th class="col-no">No.</th><th class="col-item">品名</th><th class="col-qty">数量</th><th class="col-unit">単価</th><th class="col-amount">金額</th></tr>
                </thead>
                <tbody>${itemsHtml}${emptyRowsHtml}</tbody>
            </table>
            <div class="remarks">
                <div class="remarks-title">備考</div>
                <p>有効期限: 本日より2週間<br>※本見積書はシステムによる自動発行です。</p>
            </div>
        </div>
        <script>
            window.onload = function() { setTimeout(() => { window.print(); }, 500); };
        </script>
    </body>
    </html>
    `;

    const win = window.open('', '_blank');
    win.document.write(htmlContent);
    win.document.close();
}
