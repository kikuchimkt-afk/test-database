// Global data variable
let enhancedData = [];

/**
 * Logic to estimate category (Grade, Subject, Special) from title
 */
function estimateTagsFromTitle(title) {
    if (!title) return { level: 'unknown', subject: 'unknown', special: [] };

    let level = 'unknown';
    let subject = 'unknown';
    let special = []; // eiken, exam, season, none

    // Normalize
    const t = title.toLowerCase();

    // Estimate Level
    if (t.includes('小') || t.includes('小学')) level = 'elementary';
    if (t.includes('中') || t.includes('中学') || t.includes('高校入試')) level = 'junior';
    if (t.includes('高') || t.includes('高校') || t.includes('大学入試') || t.includes('センター') || t.includes('共通テスト')) level = 'high';

    // Estimate Subject
    if (t.includes('英')) subject = 'english';
    if (t.includes('数') || t.includes('算')) subject = 'math';
    if (t.includes('国') || t.includes('文') || t.includes('漢字')) subject = 'japanese';
    if (t.includes('理') || t.includes('物') || t.includes('化') || t.includes('生')) subject = 'science';
    if (t.includes('社') || t.includes('史') || t.includes('地') || t.includes('公')) subject = 'social';

    // Estimate Special Categories
    if (t.includes('英検')) special.push('eiken');
    if (t.includes('受験') || t.includes('入試') || t.includes('共通テスト') || t.includes('センター') || t.includes('赤本') || t.includes('虎の巻')) special.push('exam');
    if (t.includes('春期') || t.includes('夏期') || t.includes('冬期') || t.includes('講習')) special.push('season');

    return { level, subject, special };
}

// === APP STATE ===
let currentState = {
    search: '',
    filterLevel: 'all',
    filterSubject: 'all',
    filterSpecial: 'all',
    sort: 'relevance',

    // Student & Cart State
    students: [], // { id, name, grade, cart: [] }
    currentStudentId: null,

    // New features
    favorites: [], // [itemId...]
    history: [], // [{date, type, detail}...]
    settings: {
        schoolName: 'ECCベストワン藍住・北島中央',
        taxRate: 0.10
    },
    viewMode: 'search' // 'search', 'favorites', 'history'
};

// Load State from LocalStorage
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
            // Always reset view filters on load/reload to ensure full list is shown initially
            currentState.search = '';
            currentState.filterLevel = 'all';
            currentState.filterSubject = 'all';
            currentState.filterSpecial = 'all';
            currentState.viewMode = 'search';
        } catch (e) { console.error('Load state failed', e); }
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

// === DOM ELEMENTS ===
const grid = document.getElementById('materialsGrid');
const searchInput = document.getElementById('searchInput');
const resultsCount = document.querySelector('#resultsCount span');
const sortSelect = document.getElementById('sortSelect');
const modalOverlay = document.getElementById('detailModal');
const closeModal = document.querySelector('.close-modal');

// Student & Cart DOM
const studentSelect = document.getElementById('studentSelect');
const addStudentBtn = document.getElementById('addStudentBtn');
const studentModal = document.getElementById('studentModal');
const saveStudentBtn = document.getElementById('saveStudentBtn');
const newStudentName = document.getElementById('newStudentName');
const cartItemsList = document.getElementById('cartItemsList');
const cartEmptyState = document.getElementById('cartEmptyState');
const cartTotalWholesale = document.getElementById('cartTotalWholesale');
const cartTotalRetail = document.getElementById('cartTotalRetail');
const exportStudentListBtn = document.getElementById('exportStudentListBtn');
const exportOrderSheetBtn = document.getElementById('exportOrderSheetBtn');
const createQuoteBtn = document.getElementById('createQuoteBtn');

// Student Mgmt extended
const newStudentGrade = document.getElementById('newStudentGrade');
const importStudentBtn = document.getElementById('importStudentBtn');
const importStudentInput = document.getElementById('importStudentInput');
const deleteAllStudentsBtn = document.getElementById('deleteAllStudentsBtn');
const deleteStudentBtn = document.getElementById('deleteStudentBtn');

// === EVENT LISTENERS ===

// Init
// Init
window.addEventListener('DOMContentLoaded', () => {
    // 1. Initialize Data
    const dataSource = (typeof DB_DATA !== 'undefined') ? DB_DATA : [];
    enhancedData = dataSource.map(item => {
        const tags = estimateTagsFromTitle(item.title);
        return { ...item, ...tags };
    });

    // 2. Load State & Reset View for clean start
    loadState();
    currentState.search = '';
    currentState.filterLevel = 'all';
    currentState.filterSubject = 'all';
    currentState.filterSpecial = 'all';
    currentState.viewMode = 'search';

    // 3. Render Initial State
    if (searchInput) searchInput.value = '';
    renderStudentSelect();
    updateCartDisplay();
    // Render grid immediately
    renderGrid();

    // Enable global export buttons initially
    if (exportStudentListBtn) exportStudentListBtn.disabled = false;
    if (exportOrderSheetBtn) exportOrderSheetBtn.disabled = false;

    // 4. Attach Listeners (with safety checks)
    const navSearch = document.getElementById('navSearch');
    if (navSearch) navSearch.addEventListener('click', () => updateViewMode('search'));

    const navFavorites = document.getElementById('navFavorites');
    if (navFavorites) navFavorites.addEventListener('click', () => updateViewMode('favorites'));

    const navHistory = document.getElementById('navHistory');
    if (navHistory) navHistory.addEventListener('click', () => updateViewMode('history'));

    const navSettings = document.getElementById('navSettings');
    if (navSettings) navSettings.addEventListener('click', openSettings);

    const saveSettingsBtn = document.getElementById('saveSettingsBtn');
    if (saveSettingsBtn) saveSettingsBtn.addEventListener('click', saveSettings);

    const btnDeleteAllStudents = document.getElementById('btnDeleteAllStudents');
    if (btnDeleteAllStudents) btnDeleteAllStudents.addEventListener('click', confirmDeleteAllStudents);

    const btnFactoryReset = document.getElementById('btnFactoryReset');
    if (btnFactoryReset) btnFactoryReset.addEventListener('click', confirmFactoryReset);

    if (deleteStudentBtn) {
        deleteStudentBtn.addEventListener('click', confirmDeleteSingleStudent);
    }

    const confirmModal = document.getElementById('confirmModal');
    if (confirmModal) {
        const cancelBtn = document.getElementById('confirmCancelBtn');
        if (cancelBtn) cancelBtn.addEventListener('click', () => confirmModal.classList.remove('active'));
        confirmModal.addEventListener('click', (e) => {
            if (e.target === confirmModal) confirmModal.classList.remove('active');
        });
    }

    const settingsModal = document.getElementById('settingsModal');
    if (settingsModal) {
        settingsModal.querySelectorAll('.close-modal, .close-modal-btn').forEach(btn => {
            btn.addEventListener('click', () => settingsModal.classList.remove('active'));
        });
    }

    // Force "Active" style on All buttons
    ['levelFilters', 'subjectFilters', 'specialFilters'].forEach(id => {
        const container = document.getElementById(id);
        if (container) {
            container.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
            const allBtn = container.querySelector('[data-filter="all"]');
            if (allBtn) allBtn.classList.add('active');
        }
    });

    updateViewMode('search');
});

// ... updateViewMode ...

function openSettings() {
    document.getElementById('settingSchoolName').value = currentState.settings.schoolName || '';
    document.getElementById('settingAddress').value = currentState.settings.address || '';
    document.getElementById('settingPhone').value = currentState.settings.phone || '';
    document.getElementById('settingTaxRate').value = currentState.settings.taxRate || 0.10;

    document.getElementById('settingsModal').classList.add('active');
}

function saveSettings() {
    currentState.settings.schoolName = document.getElementById('settingSchoolName').value;
    currentState.settings.address = document.getElementById('settingAddress').value;
    currentState.settings.phone = document.getElementById('settingPhone').value;
    currentState.settings.taxRate = parseFloat(document.getElementById('settingTaxRate').value);

    saveState();
    document.getElementById('settingsModal').classList.remove('active');
    // Optional: Flash success toast? Using alert for now as per prev interaction
    alert('設定を保存しました');
}

function showConfirm(title, message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    document.getElementById('confirmTitle').innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> ${title}`;
    document.getElementById('confirmMessage').innerText = message;

    const okBtn = document.getElementById('confirmOkBtn');

    // Clear previous listener
    const newOkBtn = okBtn.cloneNode(true);
    okBtn.parentNode.replaceChild(newOkBtn, okBtn);

    newOkBtn.addEventListener('click', () => {
        onConfirm();
        modal.classList.remove('active');
    });

    modal.classList.add('active');

    // Focus Cancel Button for safety
    setTimeout(() => {
        document.getElementById('confirmCancelBtn').focus();
    }, 50);
}

function confirmDeleteAllStudents() {
    showConfirm(
        '全生徒データ削除',
        '本当に全ての生徒データ（カート・履歴含む）を削除しますか？\nこの操作は取り消せません。\n※お気に入りや設定は残ります。',
        () => {
            currentState.students = [];
            currentState.currentStudentId = null;
            currentState.history = []; // Clear history too as it relates to students
            saveState();
            renderStudentSelect();
            updateCartDisplay();
            // Refreshes view if stuck
            renderGrid();
            alert('全生徒データを削除しました');
        }
    );
}


function confirmDeleteSingleStudent() {
    if (!currentState.currentStudentId) {
        alert('削除する生徒を選択してください');
        return;
    }

    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (!student) return;

    showConfirm(
        '生徒データ削除',
        `${student.name} さんのデータを削除しますか？\n（カートの中身も削除されますが、お気に入り等は残ります）\nこの操作は取り消せません。`,
        () => {
            currentState.students = currentState.students.filter(s => s.id !== currentState.currentStudentId);
            currentState.currentStudentId = null;
            // Optionally, select the last student if exists
            if (currentState.students.length > 0) {
                currentState.currentStudentId = currentState.students[currentState.students.length - 1].id;
            }

            saveState();
            renderStudentSelect();
            updateCartDisplay();
            alert('削除しました');
        }
    );
}

function confirmFactoryReset() {
    showConfirm(
        '完全初期化',
        'アプリケーションを完全に初期化します。\n全ての設定、お気に入り、履歴が削除されます。\n本当によろしいですか？',
        () => {
            localStorage.removeItem('materio_db_state');
            location.reload();
        }
    );
}

// Filter Buttons
document.querySelectorAll('.chip').forEach(btn => {
    btn.addEventListener('click', (e) => {
        const parentId = e.target.parentElement.id;
        const filterValue = e.target.dataset.filter;

        e.target.parentElement.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
        e.target.classList.add('active');

        if (parentId === 'levelFilters') currentState.filterLevel = filterValue;
        if (parentId === 'subjectFilters') currentState.filterSubject = filterValue;
        if (parentId === 'specialFilters') currentState.filterSpecial = filterValue;

        renderGrid();
    });
});

// Search & Sort
searchInput.addEventListener('input', (e) => { currentState.search = e.target.value.toLowerCase(); renderGrid(); });
sortSelect.addEventListener('change', (e) => { currentState.sort = e.target.value; renderGrid(); });

// Modal Logic
function openModal(item) {
    document.getElementById('modalTitle').textContent = item.title;
    document.getElementById('modalWholesale').textContent = `¥${item.price_wholesale.toLocaleString()}`;
    document.getElementById('modalRetail').textContent = `¥${item.price_retail.toLocaleString()}`;

    // Badges
    const badgesContainer = document.getElementById('modalBadges');
    badgesContainer.innerHTML = '';
    badgesContainer.appendChild(createBadge({}, item.level, 'level'));
    badgesContainer.appendChild(createBadge({}, item.subject, 'subject'));
    if (item.special) item.special.forEach(sp => badgesContainer.appendChild(createBadge({}, sp, 'special')));

    // Modal Cart Button
    const footer = modalOverlay.querySelector('.modal-footer');
    // Clear previous buttons to prevent duplicates/listener issues
    footer.innerHTML = '';

    const editBtn = document.createElement('button');
    editBtn.className = 'btn-secondary';
    editBtn.textContent = '編集';

    const addBtn = document.createElement('button');
    addBtn.className = 'btn-primary';
    addBtn.innerHTML = '<i class="fa-solid fa-cart-plus"></i> カートに追加';
    addBtn.onclick = () => addToCart(item);

    footer.appendChild(editBtn);
    footer.appendChild(addBtn);

    modalOverlay.classList.add('active');
}

closeModal.addEventListener('click', () => modalOverlay.classList.remove('active'));
studentModal.querySelector('.close-modal').addEventListener('click', () => studentModal.classList.remove('active'));
modalOverlay.addEventListener('click', (e) => { if (e.target === modalOverlay) modalOverlay.classList.remove('active'); });


// === STUDENT MANAGEMENT ===

addStudentBtn.addEventListener('click', () => {
    newStudentName.value = '';
    newStudentGrade.value = '';
    studentModal.classList.add('active');
    setTimeout(() => newStudentName.focus(), 100);
});

saveStudentBtn.addEventListener('click', () => {
    const name = newStudentName.value.trim();
    const grade = newStudentGrade.value.trim();
    if (!name) return;

    addSingleStudent(name, grade);

    studentModal.classList.remove('active');
});

function addSingleStudent(name, grade) {
    const newStudent = {
        id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
        name: name,
        grade: grade,
        cart: []
    };

    currentState.students.push(newStudent);
    currentState.currentStudentId = newStudent.id;
    saveState();

    renderStudentSelect();
    updateCartDisplay();
}

studentSelect.addEventListener('change', (e) => {
    currentState.currentStudentId = e.target.value;
    saveState();
    updateCartDisplay();
});

// Import from Excel
importStudentBtn.addEventListener('click', () => {
    importStudentInput.click();
});

importStudentInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assume first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON
            // We assume format: Column A = Name, Column B = Grade (Optional)
            // Or Header: "氏名", "学年"
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Array of Arrays

            if (jsonData.length === 0) {
                alert('データが見つかりませんでした');
                return;
            }

            let importCount = 0;

            // Analyze header to find name/grade indices
            let nameIdx = 0;
            let gradeIdx = 1;

            const headerRow = jsonData[0];
            const isHeader = headerRow.some(cell => typeof cell === 'string' && (cell.includes('氏名') || cell.includes('名前')));

            let startRow = 0;
            if (isHeader) {
                startRow = 1;
                // Find column indices
                headerRow.forEach((cell, idx) => {
                    if (typeof cell !== 'string') return;
                    if (cell.includes('氏名') || cell.includes('名前')) nameIdx = idx;
                    if (cell.includes('学年')) gradeIdx = idx;
                });
            }

            for (let i = startRow; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row || row.length === 0) continue;

                const name = (row[nameIdx] || '').toString().trim();
                const grade = (row[gradeIdx] || '').toString().trim();

                if (name) {
                    currentState.students.push({
                        // Unique ID generation
                        id: Date.now().toString() + Math.random().toString(36).substr(2, 5) + i,
                        name: name,
                        grade: grade,
                        cart: []
                    });
                    importCount++;
                }
            }

            if (importCount > 0) {
                currentState.currentStudentId = currentState.students[currentState.students.length - 1].id; // Select last imported
                saveState();
                renderStudentSelect();
                updateCartDisplay();
                alert(`${importCount}件の生徒データをインポートしました`);
            } else {
                alert('インポートできるデータがありませんでした');
            }

        } catch (error) {
            console.error(error);
            alert('ファイルの読み込みに失敗しました。\nExcelファイル形式を確認してください。');
        }
        // Reset input
        importStudentInput.value = '';
    };
    reader.readAsArrayBuffer(file);
});

// Delete All Carts (Reset)
deleteAllStudentsBtn.addEventListener('click', () => {
    if (currentState.students.length === 0) return;
    if (confirm('全ての生徒のカートを空にします。\n生徒データ自体は削除されません。\nよろしいですか？')) {
        currentState.students.forEach(student => {
            student.cart = [];
        });
        saveState();
        updateCartDisplay();
        alert('全てのカートを空にしました');
    }
});

function renderStudentSelect() {
    studentSelect.innerHTML = '<option value="">生徒を選択...</option>';
    currentState.students.forEach(student => {
        const option = document.createElement('option');
        option.value = student.id;
        // Show Name and Grade if available
        let label = student.name;
        if (student.grade) {
            label += ` (${student.grade})`;
        }
        option.textContent = label;
        if (student.id === currentState.currentStudentId) option.selected = true;
        studentSelect.appendChild(option);
    });
}


// === CART LOGIC ===

function addToCart(item) {
    if (!currentState.currentStudentId) {
        alert('先に生徒を選択してください');
        // Flash the student selector
        studentSelect.style.borderColor = 'red';
        setTimeout(() => studentSelect.style.borderColor = '', 1000);
        return;
    }

    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (!student) return;

    // Check if already in cart
    if (student.cart.some(cartItem => cartItem.id === item.id)) {
        alert('この教材は既にカートに入っています');
        return;
    }

    student.cart.push(item);
    saveState();
    updateCartDisplay();

    // Visual feedback
    // Could add toast here
    modalOverlay.classList.remove('active');
}

function removeFromCart(itemId) {
    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (!student) return;

    student.cart = student.cart.filter(item => item.id !== itemId);
    saveState();
    updateCartDisplay();
}

function updateCartDisplay() {
    cartItemsList.innerHTML = '';

    // Check selection
    if (!currentState.currentStudentId) {
        cartEmptyState.style.display = 'flex';
        cartEmptyState.querySelector('p').innerHTML = '生徒を選択して<br>教材を追加してください';
        exportCartBtn.disabled = true;
        createQuoteBtn.disabled = true;
        updateTotals(0, 0);
        return;
    }

    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (!student || student.cart.length === 0) {
        cartEmptyState.style.display = 'flex';
        cartEmptyState.querySelector('p').innerHTML = 'カートは空です';
        // Global export buttons should remain active
        // if (exportStudentListBtn) exportStudentListBtn.disabled = true;
        // if (exportOrderSheetBtn) exportOrderSheetBtn.disabled = true;
        createQuoteBtn.disabled = true;
        updateTotals(0, 0);
        return;
    }

    // Has items
    cartEmptyState.style.display = 'none';
    if (exportStudentListBtn) exportStudentListBtn.disabled = false;
    if (exportOrderSheetBtn) exportOrderSheetBtn.disabled = false;
    createQuoteBtn.disabled = false;

    let totalW = 0;
    let totalR = 0;

    student.cart.forEach(item => {
        totalW += (parseInt(item.price_wholesale) || 0);
        totalR += (parseInt(item.price_retail) || 0);

        const el = document.createElement('div');
        el.className = 'cart-item';
        el.innerHTML = `
            <div class="cart-item-title">${item.title}</div>
            <div class="cart-item-actions">
                <div class="cart-item-price">¥${(parseInt(item.price_retail) || 0).toLocaleString()}</div>
            </div>
            <button class="item-delete-btn" onclick="return false;"><i class="fa-solid fa-trash"></i></button>
        `;

        // Add delete listener
        el.querySelector('.item-delete-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            removeFromCart(item.id);
        });

        cartItemsList.appendChild(el);
    });

    updateTotals(totalW, totalR);
}

function updateTotals(wholesale, retail) {
    cartTotalWholesale.textContent = `¥${wholesale.toLocaleString()}`;
    cartTotalRetail.textContent = `¥${retail.toLocaleString()}`;
}

// === HELPER & RENDER ===

function createBadge(context, type, category) {
    const span = document.createElement('span');
    let className = 'badge';
    let label = type;

    if (category === 'level') {
        className += ` level-${type}`;
        label = getLevelLabel(type);
    } else if (category === 'subject') {
        className += ` subject-${type}`;
        label = getSubjectLabel(type);
    } else if (category === 'special') {
        className += ` special-${type}`;
        label = getSpecialLabel(type);
    }

    span.className = className;
    span.textContent = label;
    return span;
}

function getLevelLabel(type) {
    const labels = { 'elementary': '小学生', 'junior': '中学生', 'high': '高校生', 'unknown': 'その他' };
    return labels[type] || type;
}

function getSubjectLabel(type) {
    const labels = {
        'english': '英語', 'math': '数学/算数', 'japanese': '国語',
        'science': '理科', 'social': '社会', 'unknown': 'その他'
    };
    return labels[type] || type;
}

function getSpecialLabel(type) {
    const labels = { 'eiken': '英検', 'exam': '受験対策', 'season': '季節講習' };
    return labels[type] || type;
}

function renderGrid() {
    grid.innerHTML = '';

    // Mode Switching
    if (currentState.viewMode === 'history') {
        renderHistory();
        return;
    }

    // Filter
    let filtered = enhancedData.filter(item => {
        // Favorites Mode
        if (currentState.viewMode === 'favorites') {
            if (!currentState.favorites.includes(item.id)) return false;
        }

        // Search & Filter Logic
        // AND Search Implementation
        const searchKeywords = currentState.search.toLowerCase().split(/[\s\u3000]+/).filter(k => k);
        const matchSearch = searchKeywords.every(keyword => item.title.toLowerCase().includes(keyword));

        const matchLevel = currentState.filterLevel === 'all' || item.level === currentState.filterLevel;
        const matchSubject = currentState.filterSubject === 'all' || item.subject === currentState.filterSubject;

        let matchSpecial = true;
        if (currentState.filterSpecial !== 'all') {
            matchSpecial = item.special && item.special.includes(currentState.filterSpecial);
        }

        return matchSearch && matchLevel && matchSubject && matchSpecial;
    });

    // Sort
    if (currentState.sort === 'price-asc') filtered.sort((a, b) => a.price_retail - b.price_retail);
    else if (currentState.sort === 'price-desc') filtered.sort((a, b) => b.price_retail - a.price_retail);

    resultsCount.textContent = filtered.length;

    // Update count label based on view
    const countLabelBase = currentState.viewMode === 'favorites' ? 'お気に入り' : '検索結果';
    document.querySelector('#resultsCount').innerHTML = `${countLabelBase}: <span class="count-animate">${filtered.length}</span> 件`;

    // RenderItems
    filtered.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = 'material-card';
        if (index < 30) card.style.animationDelay = `${index * 0.03}s`;

        const levelLabel = getLevelLabel(item.level);
        const subjectLabel = getSubjectLabel(item.subject);

        let specialBadges = '';
        if (item.special) {
            item.special.forEach(sp => {
                specialBadges += `<span class="badge special-${sp}">${getSpecialLabel(sp)}</span>`;
            });
        }

        const isFav = currentState.favorites.includes(item.id);
        const starClass = isFav ? 'fa-solid' : 'fa-regular';
        const starColor = isFav ? '#fbbf24' : '#ccc';

        card.innerHTML = `
            <div class="card-badges">
                <span class="badge level-${item.level}">${levelLabel}</span>
                <span class="badge subject-${item.subject}">${subjectLabel}</span>
                ${specialBadges}
                <i class="${starClass} fa-star fav-btn" style="margin-left: auto; cursor: pointer; font-size: 1.2rem; color: ${starColor};" title="お気に入り"></i>
            </div>
            <div class="card-content">
                <h3>${item.title}</h3>
            </div>
            <div class="price-block">
                <div class="price-row wholesale">
                    <span class="price-label">仕入</span>
                    <span class="price-value">¥${(parseInt(item.price_wholesale) || 0).toLocaleString()}</span>
                </div>
                <div class="price-row retail">
                    <span class="price-label">販売</span>
                    <span class="price-value">¥${(parseInt(item.price_retail) || 0).toLocaleString()}</span>
                </div>
            </div>
            <button class="add-cart-btn">
                <i class="fa-solid fa-plus"></i> カートに追加
            </button>
        `;

        // Fav Toggle
        card.querySelector('.fav-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            toggleFavorite(item.id);
        });

        // Button Click -> Add to Cart directly (new feature)
        card.querySelector('.add-cart-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            addToCart(item);
        });

        grid.appendChild(card);
    });
}

function toggleFavorite(itemId) {
    const idx = currentState.favorites.indexOf(itemId);
    if (idx > -1) {
        currentState.favorites.splice(idx, 1);
    } else {
        currentState.favorites.push(itemId);
    }
    saveState();
    // Re-render if in favorites view to update list, or if just toggling star
    renderGrid();
}

function renderHistory() {
    grid.innerHTML = '';
    document.querySelector('#resultsCount').innerHTML = `履歴: <span class="count-animate">${currentState.history.length}</span> 件`;

    if (currentState.history.length === 0) {
        grid.innerHTML = '<div style="grid-column: 1/-1; text-align: center; padding: 2rem; color: #666;">履歴はありません</div>';
        return;
    }

    const list = document.createElement('div');
    list.style.gridColumn = '1 / -1';
    list.style.background = 'white';
    list.style.padding = '1rem';
    list.style.borderRadius = '8px';

    currentState.history.slice().reverse().forEach(h => {
        const item = document.createElement('div');
        item.style.borderBottom = '1px solid #eee';
        item.style.padding = '0.5rem 0';
        item.innerHTML = `
            <div><strong>${h.date}</strong> <span class="badge level-elementary">${h.type}</span></div>
            <div style="font-size: 0.9em; color: #555;">${h.detail}</div>
        `;
        list.appendChild(item);
    });
    grid.appendChild(list);
}

function addToHistory(type, detail) {
    const now = new Date();
    const dateStr = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()} ${now.getHours()}:${now.getMinutes()}`;
    currentState.history.push({
        date: dateStr,
        type: type,
        detail: detail
    });
    // Keep history manageable
    if (currentState.history.length > 50) currentState.history.shift();
    saveState();
}


// === EXPORT / QUOTATION LOGIC ===

// Quotation (PDF Print)
createQuoteBtn.addEventListener('click', () => {
    if (!currentState.currentStudentId) return;
    const student = currentState.students.find(s => s.id === currentState.currentStudentId);
    if (!student || student.cart.length === 0) return;

    printQuotation(student);
});

// A. Student List Export (ALL Students)
if (exportStudentListBtn) {
    exportStudentListBtn.addEventListener('click', () => {
        const students = currentState.students;
        if (!students || students.length === 0) {
            alert('生徒データがありません');
            return;
        }

        // Prepare Data: Name, Item Title, Retail Price (Across ALL students)
        let allData = [];
        students.forEach(student => {
            if (student.cart && student.cart.length > 0) {
                student.cart.forEach(item => {
                    allData.push({
                        '氏名': student.name,
                        '品名': item.title,
                        '価格（税込）': parseInt(item.price_retail) || 0
                    });
                });
            }
        });

        if (allData.length === 0) {
            alert('出力対象のデータがありません（カートが空です）');
            return;
        }

        // Create Worksheet
        const ws = XLSX.utils.json_to_sheet(allData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "全生徒リスト");

        // Export
        const filename = `全生徒教材リスト_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, filename);

        addToHistory('Excel出力', `全生徒教材リストを出力しました`);
    });
}

// B. Order Sheet (Summary) Export (ALL Students)
if (exportOrderSheetBtn) {
    exportOrderSheetBtn.addEventListener('click', () => {
        const students = currentState.students;
        if (!students || students.length === 0) {
            alert('生徒データがありません');
            return;
        }

        const summary = {};

        students.forEach(student => {
            if (student.cart && student.cart.length > 0) {
                student.cart.forEach(item => {
                    if (!summary[item.title]) {
                        summary[item.title] = {
                            '品名': item.title,
                            '数量': 0
                        };
                    }
                    summary[item.title]['数量']++;
                });
            }
        });

        const data = Object.values(summary);

        if (data.length === 0) {
            alert('出力対象のデータがありません（カートが空です）');
            return;
        }

        // Create Worksheet
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "発注用集計");

        const filename = `一括発注リスト_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, filename);

        addToHistory('Excel出力', `発注用集計リストを出力しました`);
    });
}

function printQuotation(student) {
    addToHistory('見積書作成', `${student.name}様の見積書を作成しました`);
    const today = new Date();
    const dateStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;

    // Calculate totals
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

    // Fill empty rows to make it look like a standard sheet if few items
    // Reduced to 10 rows minimum as per user request to fit on one page comfortable
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

    // Tax Calculation (Inclusive Pricing)
    // Assuming the registered price is Tax Inclusive
    const taxRate = currentState.settings.taxRate || 0.10;

    // Internal Tax Calculation: Total * Rate / (1 + Rate)
    const taxAmount = Math.floor(totalAmount * taxRate / (1 + taxRate));
    const priceWithoutTax = totalAmount - taxAmount;

    const htmlContent = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <title>御見積書 - ${student.name}様</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;600&display=swap');
            
            body {
                font-family: 'Noto Serif JP', serif;
                margin: 0;
                padding: 0;
                background: #ccc;
                -webkit-print-color-adjust: exact;
            }
            .page {
                width: 210mm;
                /* min-height: 297mm; removed to allow better fit check */
                height: 297mm; /* Force A4 height */
                padding: 15mm; /* reduced padding */
                margin: 10mm auto;
                background: white;
                box-shadow: 0 0 5px rgba(0,0,0,0.1);
                box-sizing: border-box;
                position: relative;
                overflow: hidden; /* Ensure no spill over */
            }
            @media print {
                body { background: none; }
                .page { margin: 0; box-shadow: none; width: 100%; height: 100%; page-break-after: always; }
                @page { margin: 0; size: A4 portrait; }
            }

            /* Header */
            .header {
                display: flex;
                justify-content: space-between;
                margin-bottom: 25px; /* Reduced */
            }
            .title {
                font-size: 20pt; /* Reduced */
                font-weight: 600;
                letter-spacing: 5px;
                border-bottom: 3px double #333;
                padding-bottom: 5px;
                display: inline-block;
            }
            .date {
                text-align: right;
                font-size: 9pt;
                margin-bottom: 5px;
            }

            /* Info Block */
            .info-block {
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                margin-bottom: 25px; /* Reduced */
            }
            .recipient {
                font-size: 14pt; /* Reduced */
                border-bottom: 1px solid #333;
                padding-bottom: 3px;
                min-width: 250px;
            }
            .recipient span {
                font-size: 10pt;
            }
            .sender {
                font-size: 9pt;
                line-height: 1.4;
                text-align: right;
            }
            .company-name {
                font-size: 11pt;
                font-weight: 600;
                margin-bottom: 3px;
            }

            /* Total Block */
            .total-block {
                margin-bottom: 20px;
                border-bottom: 2px solid #333;
                padding-bottom: 5px;
            }
            .total-label {
                font-size: 11pt;
                margin-right: 15px;
            }
            .total-amount {
                font-size: 18pt; /* Reduced */
                font-weight: 600;
            }

            /* Table */
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                font-size: 10pt;
            }
            th {
                background-color: #f0f0f0;
                border: 1px solid #333;
                padding: 6px;
                font-weight: 600;
                text-align: center;
            }
            td {
                border: 1px solid #333;
                padding: 6px;
                height: 35px; /* Fixed height for consistency */
            }
            .center { text-align: center; }
            .right { text-align: right; }
            
            .col-no { width: 30px; }
            .col-item { }
            .col-qty { width: 40px; }
            .col-unit { width: 80px; }
            .col-amount { width: 80px; }

            /* Footer / Note */
            .remarks {
                border: 1px solid #333;
                padding: 8px;
                height: 80px;
                font-size: 9pt;
            }
            .remarks-title {
                font-size: 9pt;
                font-weight: 600;
                margin-bottom: 3px;
                text-decoration: underline;
            }

        </style>
    </head>
    <body>
        <div class="page">
            <div class="date">${dateStr}</div>
            
            <div class="header">
                 <div class="title">御見積書</div>
            </div>

            <div class="info-block">
                <div class="recipient">
                    ${student.name} <span>様</span>
                </div>
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
                    <tr>
                        <th class="col-no">No.</th>
                        <th class="col-item">品名</th>
                        <th class="col-qty">数量</th>
                        <th class="col-unit">単価</th>
                        <th class="col-amount">金額</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsHtml}
                    ${emptyRowsHtml}
                </tbody>
            </table>

            <div class="remarks">
                <div class="remarks-title">備考</div>
                <p>有効期限: 本日より2週間<br>
                ※本見積書はシステムによる自動発行です。</p>
            </div>
        </div>
        <script>
            window.onload = function() {
                setTimeout(() => {
                    window.print();
                }, 500);
            };
        </script>
    </body>
    </html>
    `;

    const win = window.open('', '_blank');
    win.document.write(htmlContent);
    win.document.close();
}
