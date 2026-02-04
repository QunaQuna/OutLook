// ==================== GLOBAL VARIABLES ====================
let folders = [];
let rules = [];
let editingFolderId = null;
const MAX_FOLDER_LEVEL = 5;

// All condition types
const allConditions = [
    // People
    { id: 'from', name: 'From (từ ai)', group: 'people', param: 'email', placeholder: 'sender@company.com' },
    { id: 'sentTo', name: 'To (gửi đến ai)', group: 'people', param: 'email', placeholder: 'recipient@company.com' },
    { id: 'sentCc', name: 'Cc', group: 'people', param: 'email', placeholder: 'cc@company.com' },
    { id: 'myNameInToBox', name: "My name is on the To line", group: 'people', param: 'boolean' },
    { id: 'myNameInCcBox', name: "My name is on the Cc line", group: 'people', param: 'boolean' },
    { id: 'sentOnlyToMe', name: "I'm the only recipient", group: 'people', param: 'boolean' },
    
    // Subject
    { id: 'subjectContainsWords', name: 'Subject includes', group: 'subject', param: 'text', placeholder: 'từ khóa' },
    { id: 'subjectOrBodyContainsWords', name: 'Subject or body includes', group: 'subject', param: 'text', placeholder: 'từ khóa' },
    
    // Keywords
    { id: 'bodyContainsWords', name: 'Message body includes', group: 'keywords', param: 'text', placeholder: 'từ khóa trong nội dung' },
    { id: 'fromAddressContainsWords', name: 'Sender address includes', group: 'keywords', param: 'text', placeholder: 'phần địa chỉ email' },
    
    // Marked with
    { id: 'importance', name: 'Importance', group: 'marked', param: 'select', options: ['High', 'Normal', 'Low'] },
    
    // Message includes
    { id: 'hasAttachment', name: 'Has attachment', group: 'message', param: 'boolean' },
    { id: 'flagged', name: 'Flagged', group: 'message', param: 'boolean' },
    
    // Message size
    { id: 'messageSizeOver', name: 'Size greater than (KB)', group: 'size', param: 'number', placeholder: '1024' },
    { id: 'messageSizeUnder', name: 'Size less than (KB)', group: 'size', param: 'number', placeholder: '1024' },
    
    // Received
    { id: 'receivedAfterDate', name: 'Received after date', group: 'received', param: 'date', placeholder: 'dd/mm/yyyy' },
    { id: 'receivedBeforeDate', name: 'Received before date', group: 'received', param: 'date', placeholder: 'dd/mm/yyyy' }
];

// All action types
const allActions = [
    // Organize
    { id: 'moveToFolder', name: 'Move to folder', group: 'organize', param: 'folder' },
    { id: 'copyToFolder', name: 'Copy to folder', group: 'organize', param: 'folder' },
    { id: 'deleteMessage', name: 'Delete message', group: 'organize', param: 'none' },
    
    // Mark message
    { id: 'markAsRead', name: 'Mark as read', group: 'mark', param: 'none' },
    { id: 'markAsJunk', name: 'Mark as Junk', group: 'mark', param: 'none' },
    { id: 'setImportance', name: 'Set importance', group: 'mark', param: 'select', options: ['High', 'Normal', 'Low'] },
    { id: 'applyCategory', name: 'Apply category', group: 'mark', param: 'text', placeholder: 'Tên category' },
    
    // Route
    { id: 'forwardTo', name: 'Forward to', group: 'route', param: 'email', placeholder: 'email@company.com' },
    { id: 'redirectTo', name: 'Redirect to', group: 'route', param: 'email', placeholder: 'email@company.com' }
];

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', function() {
    console.log('Exchange Online Rules Generator v2.0 loaded');
    updateStats();
    
    // Initialize Select2
    if (typeof $.fn.select2 !== 'undefined') {
        $('.select2-folder').select2({
            placeholder: '-- Chọn folder cha --',
            allowClear: true,
            width: '100%'
        });
    }
    
    // Show welcome message
    if (typeof Swal !== 'undefined') {
        Swal.fire({
            title: 'Chào mừng!',
            html: '<p>Công cụ tạo Exchange Online Rules tự động</p><p style="font-size: 14px; color: #6c757d;">Tự động cập nhật khi thay đổi điều kiện/hành động</p>',
            icon: 'info',
            confirmButtonText: 'Bắt đầu',
            confirmButtonColor: '#667eea',
            timer: 3000,
            timerProgressBar: true
        });
    }
});

// ==================== HELPER FUNCTIONS ====================
function updateStats() {
    document.getElementById('folderCount').textContent = folders.length;
    document.getElementById('ruleCount').textContent = rules.length;
}

function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function getFolderLevel(folderId) {
    if (!folderId) return 0;
    
    const folder = folders.find(f => f.id === folderId);
    if (!folder) return 0;
    
    let level = 0;
    let current = folder;
    
    while (current && current.parentId) {
        level++;
        current = folders.find(f => f.id === current.parentId);
        if (level > MAX_FOLDER_LEVEL) break;
    }
    
    return level;
}

function canAddChildFolder(parentId) {
    const currentLevel = getFolderLevel(parentId);
    return currentLevel < MAX_FOLDER_LEVEL - 1;
}

// ==================== FOLDER FUNCTIONS ====================
function showAddFolderModal(parentId) {
    editingFolderId = null;
    
    if (parentId && !canAddChildFolder(parentId)) {
        Swal.fire({
            icon: 'warning',
            title: 'Đã đạt giới hạn',
            text: `Không thể thêm folder con (giới hạn ${MAX_FOLDER_LEVEL} cấp)`,
            confirmButtonColor: '#667eea'
        });
        return;
    }
    
    document.getElementById('modalTitle').innerHTML = '<i class="fas fa-folder-plus"></i> Thêm Folder Mới';
    document.getElementById('folderName').value = '';
    
    const parentSelect = $('#parentFolder');
    parentSelect.empty();
    parentSelect.append('<option value="">-- Folder gốc (Inbox) --</option>');
    
    folders.forEach(folder => {
        if (canAddChildFolder(folder.id)) {
            const level = getFolderLevel(folder.id);
            const indent = '　'.repeat(level);
            const option = new Option(
                indent + folder.name + ` (Cấp ${level})`,
                folder.id,
                false,
                parentId === folder.id
            );
            parentSelect.append(option);
        }
    });
    
    parentSelect.trigger('change');
    document.getElementById('folderModal').style.display = 'flex';
}

function saveFolder() {
    const name = document.getElementById('folderName').value.trim();
    const parentId = document.getElementById('parentFolder').value;
    
    if (!name) {
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Vui lòng nhập tên folder',
            confirmButtonColor: '#667eea'
        });
        return;
    }
    
    const existing = folders.find(f => 
        f.name.toLowerCase() === name.toLowerCase() && 
        f.parentId === (parentId || null)
    );
    
    if (existing) {
        Swal.fire({
            icon: 'warning',
            title: 'Trùng lặp',
            text: 'Đã có folder cùng tên trong thư mục này',
            confirmButtonColor: '#667eea'
        });
        return;
    }
    
    const newFolder = {
        id: generateId(),
        name: name,
        parentId: parentId || null
    };
    
    folders.push(newFolder);
    closeFolderModal();
    renderFolders();
    updateStats();
    
    Swal.fire({
        icon: 'success',
        title: 'Thành công!',
        text: `Đã thêm folder "${name}"`,
        timer: 1500,
        showConfirmButton: false
    });
}

function closeFolderModal() {
    document.getElementById('folderModal').style.display = 'none';
}

function addChildFolder(parentId) {
    showAddFolderModal(parentId);
}

function deleteFolder(folderId) {
    const folder = folders.find(f => f.id === folderId);
    if (!folder) return;
    
    Swal.fire({
        title: 'Xác nhận xóa',
        text: `Bạn có chắc muốn xóa folder "${folder.name}" và tất cả folder con?`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#dc3545',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Xóa',
        cancelButtonText: 'Hủy'
    }).then((result) => {
        if (result.isConfirmed) {
            const toDelete = [folderId];
            let foundNew = true;
            
            while (foundNew) {
                foundNew = false;
                folders.forEach(f => {
                    if (toDelete.includes(f.parentId) && !toDelete.includes(f.id)) {
                        toDelete.push(f.id);
                        foundNew = true;
                    }
                });
            }
            
            folders = folders.filter(f => !toDelete.includes(f.id));
            
            rules = rules.filter(rule => {
                const folderActions = rule.actions?.filter(a => 
                    a.type === 'moveToFolder' || a.type === 'copyToFolder'
                );
                if (folderActions && folderActions.some(a => toDelete.includes(a.value))) {
                    return false;
                }
                return true;
            });
            
            renderFolders();
            renderRules();
            updateStats();
            
            Swal.fire({
                icon: 'success',
                title: 'Đã xóa!',
                text: `Đã xóa ${toDelete.length} folder`,
                timer: 1500,
                showConfirmButton: false
            });
        }
    });
}

function renderFolders() {
    const container = document.getElementById('foldersTree');
    
    if (folders.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-folder-open"></i>
                <p>Chưa có folder nào. Hãy thêm folder đầu tiên!</p>
                <button class="btn btn-primary" onclick="showAddFolderModal(null)">
                    <i class="fas fa-plus"></i> Thêm Folder Gốc
                </button>
            </div>`;
        return;
    }
    
    let html = '';
    
    function renderFolderItems(folderList, level = 0) {
        folderList.forEach(folder => {
            const children = folders.filter(f => f.parentId === folder.id);
            const levelClass = level > 0 ? `level-${level}` : '';
            const canAddChild = canAddChildFolder(folder.id);
            
            html += `
                <div class="folder-item ${levelClass}">
                    <div class="folder-icon">
                        <i class="fas fa-folder"></i>
                    </div>
                    <div class="folder-name">${escapeHtml(folder.name)}</div>
                    <div class="folder-level-badge">Cấp ${level}</div>
                    <div class="folder-actions">
                        ${canAddChild ? `
                            <button class="action-btn add-child-btn" title="Thêm folder con" 
                                    onclick="addChildFolder('${folder.id}')">
                                <i class="fas fa-plus"></i>
                            </button>
                        ` : ''}
                        <button class="action-btn delete-btn" title="Xóa" 
                                onclick="deleteFolder('${folder.id}')">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </div>
            `;
            
            if (children.length > 0) {
                renderFolderItems(children, level + 1);
            }
        });
    }
    
    const rootFolders = folders.filter(f => !f.parentId);
    renderFolderItems(rootFolders);
    
    container.innerHTML = html;
}

// ==================== TEMPLATE FUNCTIONS ====================
function loadTemplate(type) {
    Swal.fire({
        title: 'Tải template?',
        text: 'Điều này sẽ xóa cấu hình hiện tại',
        icon: 'question',
        showCancelButton: true,
        confirmButtonColor: '#667eea',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Tải template',
        cancelButtonText: 'Hủy'
    }).then((result) => {
        if (result.isConfirmed) {
            folders = [];
            rules = [];
            
            if (type === 'accounting') {
                loadAccountingTemplate();
            } else if (type === 'project') {
                loadProjectTemplate();
            } else if (type === 'simple') {
                loadSimpleTemplate();
            }
            
            renderFolders();
            renderRules();
            updateStats();
            
            Swal.fire({
                icon: 'success',
                title: 'Đã tải template!',
                timer: 1500,
                showConfirmButton: false
            });
        }
    });
}

function loadAccountingTemplate() {
    const root1 = { id: generateId(), name: "01. HOA DON", parentId: null };
    folders.push(root1);
    
    const dauVao = { id: generateId(), name: "DAU VAO", parentId: root1.id };
    folders.push(dauVao);
    
    const dauRa = { id: generateId(), name: "DAU RA", parentId: root1.id };
    folders.push(dauRa);
    
    folders.push({ id: generateId(), name: "THANG 01", parentId: dauVao.id });
    folders.push({ id: generateId(), name: "THANG 02", parentId: dauVao.id });
    folders.push({ id: generateId(), name: "THANG 03", parentId: dauVao.id });
    
    const root2 = { id: generateId(), name: "02. BAO CAO", parentId: null };
    folders.push(root2);
    
    folders.push({ id: generateId(), name: "THUE", parentId: root2.id });
    folders.push({ id: generateId(), name: "TAI CHINH", parentId: root2.id });
    
    rules.push({
        id: generateId(),
        name: "Hoa don dau vao",
        conditions: [
            { id: generateId(), type: 'subjectContainsWords', value: '', keywords: ['hoa don', 'invoice'] },
            { id: generateId(), type: 'hasAttachment', value: 'true' }
        ],
        actions: [
            { type: 'moveToFolder', value: dauVao.id },
            { type: 'applyCategory', value: 'Hoa don' }
        ],
        enabled: true,
        expanded: false
    });
}

function loadProjectTemplate() {
    const duAn = { id: generateId(), name: "DU AN A", parentId: null };
    folders.push(duAn);
    
    const tailieu = { id: generateId(), name: "01. TAILIEU", parentId: duAn.id };
    folders.push(tailieu);
    
    folders.push({ id: generateId(), name: "HOP DONG", parentId: tailieu.id });
    folders.push({ id: generateId(), name: "THIET KE", parentId: tailieu.id });
    folders.push({ id: generateId(), name: "BAN VE KY THUAT", parentId: tailieu.id });
    
    const lienhe = { id: generateId(), name: "02. LIEN HE", parentId: duAn.id };
    folders.push(lienhe);
    
    folders.push({ id: generateId(), name: "KHACH HANG", parentId: lienhe.id });
    folders.push({ id: generateId(), name: "DOI TAC", parentId: lienhe.id });
    
    const contractFolder = folders.find(f => f.name === "HOP DONG");
    if (contractFolder) {
        rules.push({
            id: generateId(),
            name: "Hop dong du an",
            conditions: [
                { id: generateId(), type: 'subjectContainsWords', value: '', keywords: ['hop dong', 'contract'] }
            ],
            actions: [
                { type: 'moveToFolder', value: contractFolder.id }
            ],
            enabled: true,
            expanded: false
        });
    }
}

function loadSimpleTemplate() {
    folders.push({ id: generateId(), name: "01. CAN XU LY", parentId: null });
    folders.push({ id: generateId(), name: "02. DA XU LY", parentId: null });
    folders.push({ id: generateId(), name: "03. LUU TRU", parentId: null });
    
    const processFolder = folders[0];
    rules.push({
        id: generateId(),
        name: "Email khan cap",
        conditions: [
            { id: generateId(), type: 'importance', value: 'High' }
        ],
        actions: [
            { type: 'moveToFolder', value: processFolder.id },
            { type: 'applyCategory', value: 'Khan cap' }
        ],
        enabled: true,
        expanded: false
    });
}

function clearFolders() {
    if (folders.length === 0) {
        Swal.fire({
            icon: 'info',
            title: 'Chưa có folder nào',
            timer: 1500,
            showConfirmButton: false
        });
        return;
    }
    
    Swal.fire({
        title: 'Xóa tất cả?',
        text: 'Điều này sẽ xóa tất cả folder và rules',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#dc3545',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Xóa tất cả',
        cancelButtonText: 'Hủy'
    }).then((result) => {
        if (result.isConfirmed) {
            folders = [];
            rules = [];
            renderFolders();
            renderRules();
            updateStats();
            
            Swal.fire({
                icon: 'success',
                title: 'Đã xóa tất cả!',
                timer: 1500,
                showConfirmButton: false
            });
        }
    });
}

// ==================== RULE FUNCTIONS ====================
function addNewRule() {
    const newRule = {
        id: generateId(),
        name: `Rule ${rules.length + 1}`,
        conditions: [
            { id: generateId(), type: allConditions[0].id, value: '', keywords: [] }
        ],
        actions: [
            { type: allActions[0].id, value: allActions[0].param === 'folder' && folders.length > 0 ? folders[0].id : '' }
        ],
        enabled: true,
        expanded: true
    };
    rules.push(newRule);
    renderRules();
    updateStats();
}

function deleteRule(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    Swal.fire({
        title: 'Xác nhận xóa',
        text: `Bạn có chắc muốn xóa rule "${rule.name}"?`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#dc3545',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Xóa',
        cancelButtonText: 'Hủy'
    }).then((result) => {
        if (result.isConfirmed) {
            rules = rules.filter(r => r.id !== ruleId);
            renderRules();
            updateStats();
            
            Swal.fire({
                icon: 'success',
                title: 'Đã xóa!',
                timer: 1500,
                showConfirmButton: false
            });
        }
    });
}

function toggleRuleExpanded(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.expanded = !rule.expanded;
        renderRules();
    }
}

function toggleRuleEnabled(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.enabled = !rule.enabled;
        renderRules();
    }
}

function updateRuleName(ruleId, name) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.name = name;
    }
}

// ==================== AUTO-UPDATE CONDITION FUNCTIONS ====================
function addCondition(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    if (!rule.conditions) rule.conditions = [];
    
    rule.conditions.push({
        id: generateId(),
        type: allConditions[0].id,
        value: '',
        keywords: []
    });
    renderRules();
}

function removeCondition(ruleId, conditionId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.conditions) {
        rule.conditions = rule.conditions.filter(c => c.id !== conditionId);
        renderRules();
    }
}

// AUTO-UPDATE: Tự động cập nhật khi thay đổi type và value
function updateCondition(ruleId, conditionId, type, value) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.conditions) {
        const condition = rule.conditions.find(c => c.id === conditionId);
        if (condition) {
            const oldType = condition.type;
            condition.type = type;
            condition.value = value;
            
            // Nếu thay đổi type, reset keywords
            if (oldType !== type) {
                condition.keywords = [];
            }
            
            // Tự động render lại
            renderRules();
        }
    }
}

// ==================== AUTO-UPDATE ACTION FUNCTIONS ====================
function addAction(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    if (!rule.actions) rule.actions = [];
    
    rule.actions.push({
        type: allActions[0].id,
        value: allActions[0].param === 'folder' && folders.length > 0 ? folders[0].id : ''
    });
    renderRules();
}

function removeAction(ruleId, index) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.actions) {
        rule.actions.splice(index, 1);
        renderRules();
    }
}

// AUTO-UPDATE: Tự động cập nhật khi thay đổi type và value
function updateAction(ruleId, index, type, value) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.actions && rule.actions[index]) {
        rule.actions[index].type = type;
        rule.actions[index].value = value;
        
        // Tự động render lại
        renderRules();
    }
}

// ==================== KEYWORD FUNCTIONS ====================
function addConditionKeyword(ruleId, conditionId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    const condition = rule.conditions?.find(c => c.id === conditionId);
    if (!condition) return;
    
    if (!condition.keywords) condition.keywords = [];
    
    const trimmedKeyword = keyword.trim();
    if (!trimmedKeyword) return;
    
    if (!condition.keywords.includes(trimmedKeyword)) {
        condition.keywords.push(trimmedKeyword);
        renderRules();
    }
}

function removeConditionKeyword(ruleId, conditionId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    const condition = rule.conditions?.find(c => c.id === conditionId);
    if (condition && condition.keywords) {
        condition.keywords = condition.keywords.filter(k => k !== keyword);
        renderRules();
    }
}

function handleConditionKeywordKeydown(event, ruleId, conditionId) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const input = event.target;
        const keyword = input.value.trim();
        
        if (keyword) {
            addConditionKeyword(ruleId, conditionId, keyword);
            input.value = '';
        }
    } else if (event.key === 'Backspace' && event.target.value === '') {
        const rule = rules.find(r => r.id === ruleId);
        const condition = rule?.conditions?.find(c => c.id === conditionId);
        if (condition && condition.keywords && condition.keywords.length > 0) {
            condition.keywords.pop();
            renderRules();
        }
    }
}

function handleConditionKeywordBlur(event, ruleId, conditionId) {
    const input = event.target;
    const keyword = input.value.trim();
    
    if (keyword) {
        addConditionKeyword(ruleId, conditionId, keyword);
        input.value = '';
    }
}

// ==================== RENDER FUNCTIONS ====================
function renderRules() {
    const container = document.getElementById('rulesContainer');
    
    if (rules.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-filter"></i>
                <p>Chưa có rule nào. Hãy thêm rule đầu tiên!</p>
                <button class="btn btn-primary" onclick="addNewRule()">
                    <i class="fas fa-plus"></i> Thêm Rule
                </button>
            </div>`;
        return;
    }
    
    let html = '';
    
    rules.forEach(rule => {
        const conditionsText = renderConditionsSummary(rule);
        const actionsText = renderActionsSummary(rule);
        
        html += `
            <div class="rule-item">
                <div class="rule-header">
                    <div class="rule-controls">
                        <input type="checkbox" ${rule.enabled ? 'checked' : ''} 
                               onchange="toggleRuleEnabled('${rule.id}')"
                               title="Kích hoạt/Tắt rule">
                        <input type="text" class="rule-title-input" value="${escapeHtml(rule.name)}" 
                               onchange="updateRuleName('${rule.id}', this.value)">
                        <button class="toggle-rule-btn" onclick="toggleRuleExpanded('${rule.id}')">
                            <i class="fas fa-${rule.expanded ? 'chevron-up' : 'chevron-down'}"></i> 
                            ${rule.expanded ? 'Thu gọn' : 'Chi tiết'}
                        </button>
                    </div>
                    <button class="action-btn delete-btn" onclick="deleteRule('${rule.id}')" title="Xóa rule">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
                
                <div class="conditions-summary">
                    <strong><i class="fas fa-filter"></i> Điều kiện:</strong> 
                    ${conditionsText || '<span style="color: #6c757d; font-style: italic;">Chưa có điều kiện</span>'}
                </div>
                
                <div class="actions-summary">
                    <strong><i class="fas fa-bolt"></i> Hành động:</strong> 
                    ${actionsText || '<span style="color: #6c757d; font-style: italic;">Chưa có hành động</span>'}
                </div>
                
                ${rule.expanded ? renderRuleExpanded(rule) : ''}
            </div>
        `;
    });
    
    container.innerHTML = html;
}

function renderConditionsSummary(rule) {
    if (!rule.conditions || rule.conditions.length === 0) return '';
    
    return rule.conditions.map(cond => {
        const conditionDef = allConditions.find(c => c.id === cond.type);
        if (!conditionDef) return '';
        
        let displayValue = '';
        
        if (cond.keywords && cond.keywords.length > 0) {
            displayValue = cond.keywords.join(', ');
        } else if (cond.value) {
            if (conditionDef.param === 'boolean') {
                displayValue = cond.value === 'true' ? 'Có' : 'Không';
            } else if (conditionDef.param === 'folder') {
                const folder = folders.find(f => f.id === cond.value);
                displayValue = folder ? folder.name : '(chọn folder)';
            } else {
                displayValue = cond.value;
            }
        }
        
        return `${conditionDef.name}${displayValue ? ': ' + displayValue : ''}`;
    }).filter(text => text).join(' • ');
}

function renderActionsSummary(rule) {
    if (!rule.actions || rule.actions.length === 0) return '';
    
    return rule.actions.map(action => {
        const actionDef = allActions.find(a => a.id === action.type);
        if (!actionDef) return '';
        
        let displayValue = action.value;
        if (actionDef.param === 'folder') {
            const folder = folders.find(f => f.id === action.value);
            displayValue = folder ? folder.name : '(chọn folder)';
        }
        
        return `${actionDef.name}${displayValue ? ': ' + displayValue : ''}`;
    }).filter(text => text).join(' • ');
}

function renderRuleExpanded(rule) {
    return `
        <div class="rule-expanded">
            <div class="condition-group">
                <div class="group-title">
                    <i class="fas fa-filter"></i> Điều kiện (tự động cập nhật)
                </div>
                ${rule.conditions && rule.conditions.length > 0 ? 
                    rule.conditions.map((condition, index) => renderCondition(rule.id, index, condition)).join('') : 
                    '<div style="color: #6c757d; font-style: italic; padding: 10px;">Chưa có điều kiện nào</div>'}
                <button class="add-btn" onclick="addCondition('${rule.id}')">
                    <i class="fas fa-plus"></i> Thêm điều kiện
                </button>
            </div>
            
            <div class="action-group">
                <div class="group-title">
                    <i class="fas fa-bolt"></i> Hành động (tự động cập nhật)
                </div>
                ${rule.actions && rule.actions.length > 0 ? 
                    rule.actions.map((action, index) => renderAction(rule.id, index, action)).join('') : 
                    '<div style="color: #6c757d; font-style: italic; padding: 10px;">Chưa có hành động nào</div>'}
                <button class="add-btn" onclick="addAction('${rule.id}')">
                    <i class="fas fa-plus"></i> Thêm hành động
                </button>
            </div>
        </div>
    `;
}

function renderCondition(ruleId, index, condition) {
    const conditionDef = allConditions.find(c => c.id === condition.type) || allConditions[0];
    const inputHtml = renderConditionInput(ruleId, condition.id, condition, conditionDef);
    
    return `
        <div class="condition-row">
            <select class="condition-type-select" onchange="updateCondition('${ruleId}', '${condition.id}', this.value, '')">
                ${allConditions.map(c => 
                    `<option value="${c.id}" ${condition.type === c.id ? 'selected' : ''}>${c.name}</option>`
                ).join('')}
            </select>
            ${inputHtml}
            <button class="remove-btn" onclick="removeCondition('${ruleId}', '${condition.id}')" title="Xóa điều kiện">
                <i class="fas fa-times"></i>
            </button>
        </div>
    `;
}

function renderConditionInput(ruleId, conditionId, condition, conditionDef) {
    const supportsTags = conditionDef.param === 'text';
    
    if (supportsTags) {
        return renderConditionTagInput(ruleId, conditionId, condition, conditionDef);
    }
    
    switch(conditionDef.param) {
        case 'select':
            return `
                <select class="condition-value-select" onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)">
                    ${conditionDef.options.map(opt => 
                        `<option value="${opt}" ${condition.value === opt ? 'selected' : ''}>${opt}</option>`
                    ).join('')}
                </select>
            `;
        case 'folder':
            return `
                <select class="condition-value-select" onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)">
                    <option value="">-- Chọn folder --</option>
                    ${folders.map(folder => 
                        `<option value="${folder.id}" ${condition.value === folder.id ? 'selected' : ''}>${getFolderPath(folder.id)}</option>`
                    ).join('')}
                </select>
            `;
        case 'boolean':
            return `
                <select class="condition-value-select" onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)">
                    <option value="true" ${condition.value === 'true' ? 'selected' : ''}>Có</option>
                    <option value="false" ${condition.value === 'false' ? 'selected' : ''}>Không</option>
                </select>
            `;
        case 'number':
            return `
                <input type="number" class="condition-value-input" value="${condition.value || ''}" 
                       onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)"
                       placeholder="${conditionDef.placeholder}">
                <span>KB</span>
            `;
        case 'date':
        case 'email':
            return `
                <input type="${conditionDef.param}" class="condition-value-input" value="${condition.value || ''}" 
                       onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)"
                       placeholder="${conditionDef.placeholder}">
            `;
        case 'none':
            return `<div class="condition-value-input" style="color: #6c757d; font-style: italic;">Không cần giá trị</div>`;
        default:
            return `
                <input type="text" class="condition-value-input" value="${condition.value || ''}" 
                       onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)"
                       placeholder="${conditionDef.placeholder}">
            `;
    }
}

function renderConditionTagInput(ruleId, conditionId, condition, conditionDef) {
    const keywords = condition.keywords || [];
    
    return `
        <div class="condition-tags-container" style="flex: 1; min-width: 200px;">
            <div class="keywords-input-container">
                ${keywords.map(keyword => `
                    <div class="keyword-tag">
                        ${escapeHtml(keyword)}
                        <button type="button" class="keyword-tag-remove" 
                                onclick="removeConditionKeyword('${ruleId}', '${conditionId}', '${escapeHtml(keyword)}')">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                `).join('')}
                <input type="text" 
                       class="keywords-input" 
                       placeholder="${conditionDef.placeholder}... (Enter để thêm)"
                       onkeydown="handleConditionKeywordKeydown(event, '${ruleId}', '${conditionId}')"
                       onblur="handleConditionKeywordBlur(event, '${ruleId}', '${conditionId}')">
            </div>
        </div>
    `;
}

function renderAction(ruleId, index, action) {
    const actionDef = allActions.find(a => a.id === action.type) || allActions[0];
    const inputHtml = renderActionInput(ruleId, index, action, actionDef);
    
    return `
        <div class="action-row">
            <select class="action-type-select" onchange="updateAction('${ruleId}', ${index}, this.value, '')">
                ${allActions.map(a => 
                    `<option value="${a.id}" ${action.type === a.id ? 'selected' : ''}>${a.name}</option>`
                ).join('')}
            </select>
            ${inputHtml}
            <button class="remove-btn" onclick="removeAction('${ruleId}', ${index})" title="Xóa hành động">
                <i class="fas fa-times"></i>
            </button>
        </div>
    `;
}

function renderActionInput(ruleId, index, action, actionDef) {
    switch(actionDef.param) {
        case 'select':
            return `
                <select class="action-value-select" onchange="updateAction('${ruleId}', ${index}, '${action.type}', this.value)">
                    ${actionDef.options.map(opt => 
                        `<option value="${opt}" ${action.value === opt ? 'selected' : ''}>${opt}</option>`
                    ).join('')}
                </select>
            `;
        case 'folder':
            return `
                <select class="action-value-select" onchange="updateAction('${ruleId}', ${index}, '${action.type}', this.value)">
                    <option value="">-- Chọn folder --</option>
                    ${folders.map(folder => 
                        `<option value="${folder.id}" ${action.value === folder.id ? 'selected' : ''}>${getFolderPath(folder.id)}</option>`
                    ).join('')}
                </select>
            `;
        case 'email':
        case 'text':
            return `
                <input type="${actionDef.param === 'email' ? 'email' : 'text'}" 
                       class="action-value-input" value="${action.value || ''}" 
                       onchange="updateAction('${ruleId}', ${index}, '${action.type}', this.value)"
                       placeholder="${actionDef.placeholder}">
            `;
        case 'none':
            return `<div class="action-value-input" style="color: #6c757d; font-style: italic;">Không cần giá trị</div>`;
        default:
            return `
                <input type="text" class="action-value-input" value="${action.value || ''}" 
                       onchange="updateAction('${ruleId}', ${index}, '${action.type}', this.value)"
                       placeholder="${actionDef.placeholder}">
            `;
    }
}

function getFolderPath(folderId) {
    const folder = folders.find(f => f.id === folderId);
    if (!folder) return '';
    
    let path = folder.name;
    let current = folder;
    
    while (current && current.parentId) {
        const parent = folders.find(f => f.id === current.parentId);
        if (parent) {
            path = parent.name + ' → ' + path;
            current = parent;
        } else {
            break;
        }
    }
    
    return path;
}

function getFolderPathForScript(folderId) {
    const folder = folders.find(f => f.id === folderId);
    if (!folder) return '';
    
    let path = folder.name;
    let current = folder;
    
    while (current && current.parentId) {
        const parent = folders.find(f => f.id === current.parentId);
        if (parent) {
            path = parent.name + '\\' + path;
            current = parent;
        } else {
            break;
        }
    }
    
    return path;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function escapePS(text) {
    if (!text) return '';
    // Double quotes for PowerShell string escaping
    // Remove problematic characters that might break PowerShell
    return text
        .replace(/"/g, '""')
        .replace(/\$/g, '`$')
        .replace(/`/g, '``');
}

// ==================== SCRIPT GENERATION ====================
function generateScriptContent(userEmail, scriptName, additionalNotes) {
    const createAllFolders = document.getElementById('createAllFolders').checked;
    const enableAllRules = document.getElementById('enableAllRules').checked;
    
    if (!scriptName.toLowerCase().endsWith('.ps1')) {
        scriptName += '.ps1';
    }
    
    let script = `# ${scriptName}\n`;
    script += `# Outlook Rules Script - Tao tu dong v2.0\n`;
    script += `# Email: ${userEmail}\n`;
    script += `# Ngay tao: ${new Date().toLocaleDateString('vi-VN')}\n`;
    
    if (additionalNotes) {
        script += `# Ghi chu: ${additionalNotes}\n`;
    }
    
    script += `\n# ==================== CAU HINH ====================\n`;
    script += `# Cau hinh UTF-8 cho tieng Viet\n`;
    script += `try {\n`;
    script += `    chcp 65001 | Out-Null\n`;
    script += `    \$OutputEncoding = [System.Text.Encoding]::UTF8\n`;
    script += `} catch {\n`;
    script += `    Write-Host "Warning: Could not set UTF-8 encoding" -ForegroundColor Yellow\n`;
    script += `}\n\n`;
    
    script += `Write-Host "========================================================" -ForegroundColor Cyan\n`;
    script += `Write-Host "     TAO OUTLOOK RULES TU DONG                          " -ForegroundColor Cyan\n`;
    script += `Write-Host "========================================================" -ForegroundColor Cyan\n`;
    script += `Write-Host "Email: ${userEmail}" -ForegroundColor Yellow\n`;
    script += `Write-Host ""\n\n`;
    
    script += `# Ket noi Exchange Online\n`;
    script += `try {\n`;
    script += `    Connect-ExchangeOnline -UserPrincipalName "${userEmail}" -ShowProgress \$true\n`;
    script += `    Write-Host "OK - Da ket noi Exchange Online" -ForegroundColor Green\n`;
    script += `} catch {\n`;
    script += `    Write-Host "LOI - Loi ket noi Exchange Online" -ForegroundColor Red\n`;
    script += `    Write-Host "Chi tiet: \$_" -ForegroundColor Red\n`;
    script += `    exit 1\n`;
    script += `}\n`;
    script += `\$mailbox = "${userEmail}"\n\n`;
    
    // Tạo folders
    if (folders.length > 0 && createAllFolders) {
        script += `# ==================== TAO FOLDERS (${folders.length} folders) ====================\n`;
        script += `Write-Host "Tao ${folders.length} folder..." -ForegroundColor Cyan\n`;
        script += `\$folderPaths = @{}\n\n`;
        
        const sortedFolders = [...folders].sort((a, b) => {
            return getFolderLevel(a.id) - getFolderLevel(b.id);
        });
        
        sortedFolders.forEach(folder => {
            const level = getFolderLevel(folder.id);
            const escapedName = escapePS(folder.name);
            
            script += `# Level ${level}: ${folder.name}\n`;
            script += `try {\n`;
            
            if (!folder.parentId) {
                script += `    New-MailboxFolder -Name "${escapedName}" -Parent "\$mailbox\`:\\Inbox" -ErrorAction SilentlyContinue\n`;
                script += `    \$folderPaths['${folder.id}'] = "\$mailbox\`:\\Inbox\\${escapedName}"\n`;
            } else {
                script += `    \$parentPath = \$folderPaths['${folder.parentId}']\n`;
                script += `    if (\$parentPath) {\n`;
                script += `        New-MailboxFolder -Name "${escapedName}" -Parent "\$parentPath" -ErrorAction SilentlyContinue\n`;
                script += `        \$folderPaths['${folder.id}'] = "\$parentPath\\${escapedName}"\n`;
                script += `    }\n`;
            }
            
            script += `    Write-Host "  OK - ${folder.name}" -ForegroundColor Green\n`;
            script += `} catch {\n`;
            script += `    Write-Host "  WARNING - ${folder.name} (co the da ton tai)" -ForegroundColor Yellow\n`;
            script += `}\n\n`;
        });
    }
    
    // Tạo rules
    const enabledRules = rules.filter(r => r.enabled && enableAllRules);
    if (enabledRules.length > 0) {
        script += `# ==================== TAO RULES (${enabledRules.length} rules) ====================\n`;
        script += `Write-Host "Tao ${enabledRules.length} rules..." -ForegroundColor Cyan\n\n`;
        
        enabledRules.forEach((rule, index) => {
            script += `# Rule ${index + 1}: ${rule.name}\n`;
            script += `try {\n`;
            script += `    Write-Host "Rule ${index + 1}: ${rule.name}" -NoNewline\n`;
            
            let cmdParams = [];
            cmdParams.push(`-Name "${escapePS(rule.name)}"`);
            
            // Add conditions
            if (rule.conditions && rule.conditions.length > 0) {
                rule.conditions.forEach(condition => {
                    const conditionDef = allConditions.find(c => c.id === condition.type);
                    
                    if (conditionDef) {
                        if (condition.keywords && condition.keywords.length > 0) {
                            const keywordsString = condition.keywords.map(k => `"${escapePS(k)}"`).join(', ');
                            
                            switch(condition.type) {
                                case 'subjectContainsWords':
                                    cmdParams.push(`-SubjectContainsWords ${keywordsString}`);
                                    break;
                                case 'subjectOrBodyContainsWords':
                                    cmdParams.push(`-SubjectOrBodyContainsWords ${keywordsString}`);
                                    break;
                                case 'bodyContainsWords':
                                    cmdParams.push(`-BodyContainsWords ${keywordsString}`);
                                    break;
                                case 'fromAddressContainsWords':
                                    cmdParams.push(`-FromAddressContainsWords ${keywordsString}`);
                                    break;
                            }
                        } else if (condition.value && condition.value.trim() !== '') {
                            switch(condition.type) {
                                case 'from':
                                    cmdParams.push(`-From "${escapePS(condition.value)}"`);
                                    break;
                                case 'sentTo':
                                    cmdParams.push(`-SentTo "${escapePS(condition.value)}"`);
                                    break;
                                case 'importance':
                                    cmdParams.push(`-WithImportance "${condition.value}"`);
                                    break;
                                case 'hasAttachment':
                                    if (condition.value === 'true') {
                                        cmdParams.push(`-HasAttachment \$true`);
                                    }
                                    break;
                                case 'messageSizeOver':
                                    cmdParams.push(`-MessageSizeOver ${condition.value}`);
                                    break;
                                case 'messageSizeUnder':
                                    cmdParams.push(`-MessageSizeUnder ${condition.value}`);
                                    break;
                            }
                        }
                    }
                });
            }
            
            // Add actions
            if (rule.actions && rule.actions.length > 0) {
                rule.actions.forEach(action => {
                    const actionDef = allActions.find(a => a.id === action.type);
                    
                    if (actionDef && action.value) {
                        switch(action.type) {
                            case 'moveToFolder':
                                const moveFolder = folders.find(f => f.id === action.value);
                                if (moveFolder) {
                                    const folderPath = getFolderPathForScript(moveFolder.id);
                                    cmdParams.push(`-MoveToFolder "\$mailbox\`:\\Inbox\\${escapePS(folderPath)}"`);
                                }
                                break;
                            case 'copyToFolder':
                                const copyFolder = folders.find(f => f.id === action.value);
                                if (copyFolder) {
                                    const folderPath = getFolderPathForScript(copyFolder.id);
                                    cmdParams.push(`-CopyToFolder "\$mailbox\`:\\Inbox\\${escapePS(folderPath)}"`);
                                }
                                break;
                            case 'deleteMessage':
                                cmdParams.push(`-DeleteMessage \$true`);
                                break;
                            case 'markAsRead':
                                cmdParams.push(`-MarkAsRead \$true`);
                                break;
                            case 'applyCategory':
                                cmdParams.push(`-ApplyCategory "${escapePS(action.value)}"`);
                                break;
                            case 'forwardTo':
                                cmdParams.push(`-ForwardTo "${escapePS(action.value)}"`);
                                break;
                            case 'redirectTo':
                                cmdParams.push(`-RedirectTo "${escapePS(action.value)}"`);
                                break;
                        }
                    }
                });
            }
            
            const cmd = `    New-InboxRule ${cmdParams.join(' ')} -ErrorAction Stop`;
            script += cmd + `\n`;
            script += `    Write-Host " OK" -ForegroundColor Green\n`;
            script += `} catch {\n`;
            script += `    Write-Host " LOI" -ForegroundColor Red\n`;
            script += `    Write-Host "    Loi: \$_" -ForegroundColor Red\n`;
            script += `}\n\n`;
        });
    }
    
    script += `# ==================== HOAN TAT ====================\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "========================================================" -ForegroundColor Green\n`;
    script += `Write-Host "                DA HOAN THANH!                            " -ForegroundColor Green\n`;
    script += `Write-Host "========================================================" -ForegroundColor Green\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "Tong so folder: ${folders.length}" -ForegroundColor Cyan\n`;
    script += `Write-Host "Tong so rules: ${enabledRules.length}" -ForegroundColor Cyan\n`;
    script += `Write-Host ""\n`;
    script += `Disconnect-ExchangeOnline -Confirm:\$false\n`;
    script += `Write-Host "Da ngat ket noi Exchange Online" -ForegroundColor Gray\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "Nhan Enter de thoat..." -ForegroundColor Gray\n`;
    script += `pause\n`;
    
    return { script, fileName: scriptName };
}

// ==================== PREVIEW & DOWNLOAD ====================
function previewScript() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_outlook_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        Swal.fire({
            icon: 'error',
            title: 'Thiếu thông tin',
            text: 'Vui lòng nhập email tài khoản Exchange Online',
            confirmButtonColor: '#667eea'
        });
        return;
    }
    
    const { script } = generateScriptContent(userEmail, scriptName, additionalNotes);
    document.getElementById('previewText').textContent = script;
    document.getElementById('previewModal').style.display = 'flex';
}

function closePreviewModal() {
    document.getElementById('previewModal').style.display = 'none';
}

function downloadFromPreview() {
    downloadScriptFile();
}

function copyScriptToClipboard() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_outlook_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        Swal.fire({
            icon: 'error',
            title: 'Thiếu thông tin',
            text: 'Vui lòng nhập email tài khoản Exchange Online',
            confirmButtonColor: '#667eea'
        });
        return;
    }
    
    const { script } = generateScriptContent(userEmail, scriptName, additionalNotes);
    
    navigator.clipboard.writeText(script).then(() => {
        Swal.fire({
            icon: 'success',
            title: 'Đã sao chép!',
            text: 'Script đã được sao chép vào clipboard',
            timer: 1500,
            showConfirmButton: false
        });
    }).catch(() => {
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Không thể sao chép, vui lòng thử lại',
            confirmButtonColor: '#667eea'
        });
    });
}

function downloadScriptFile() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_outlook_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        Swal.fire({
            icon: 'error',
            title: 'Thiếu thông tin',
            text: 'Vui lòng nhập email tài khoản Exchange Online',
            confirmButtonColor: '#667eea'
        });
        return;
    }
    
    const { script, fileName } = generateScriptContent(userEmail, scriptName, additionalNotes);
    
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + script], { 
        type: 'text/plain;charset=utf-8' 
    });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.style.display = 'none';
    document.body.appendChild(a);
    
    a.click();
    
    setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }, 100);
    
    Swal.fire({
        icon: 'success',
        title: 'Đã tải xuống!',
        html: `File <strong>${fileName}</strong> đã được tải xuống<br><br>
               <small>Chạy script trong PowerShell với quyền Admin</small>`,
        confirmButtonColor: '#667eea'
    });
}