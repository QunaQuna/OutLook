
// ==================== GLOBAL VARIABLES ====================
let folders = [];
let rules = [];
let editingFolderId = null;

// Tất cả điều kiện trong một mảng lớn
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

// Tất cả hành động trong một mảng lớn
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
    console.log('Outlook Rules Generator loaded');
    showEmptyStates();
});

// ==================== HELPER FUNCTIONS ====================
function showEmptyStates() {
    if (folders.length === 0) {
        document.getElementById('emptyFolders').style.display = 'block';
        const foldersContainer = document.getElementById('foldersTree');
        foldersContainer.innerHTML = `
            <div class="empty-state" id="emptyFolders">
                <i class="fas fa-folder-open"></i>
                <p>Chưa có folder nào. Hãy thêm folder đầu tiên!</p>
                <button class="btn btn-primary" onclick="showAddFolderModal(null)">
                    <i class="fas fa-plus"></i> Thêm Folder Gốc
                </button>
            </div>`;
    }
    
    if (rules.length === 0) {
        document.getElementById('emptyRules').style.display = 'block';
        const rulesContainer = document.getElementById('rulesContainer');
        rulesContainer.innerHTML = `
            <div class="empty-state" id="emptyRules">
                <i class="fas fa-filter"></i>
                <p>Chưa có rule nào. Hãy thêm rule đầu tiên!</p>
                <button class="btn btn-primary" onclick="addNewRule()">
                    <i class="fas fa-plus"></i> Thêm Rule
                </button>
            </div>`;
    }
}

function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// ==================== FOLDER FUNCTIONS ====================
function showAddFolderModal(parentId) {
    editingFolderId = null;
    document.getElementById('modalTitle').textContent = 'Thêm Folder Mới';
    document.getElementById('folderName').value = '';
    
    const parentSelect = document.getElementById('parentFolder');
    parentSelect.innerHTML = '<option value="">-- Chọn folder cha (Folder gốc) --</option>';
    
    // Add inbox option
    const inboxOption = document.createElement('option');
    inboxOption.value = 'INBOX';
    inboxOption.textContent = 'Inbox (Gốc)';
    parentSelect.appendChild(inboxOption);
    
    // Add root folders
    folders.filter(f => !f.parentId).forEach(folder => {
        const option = document.createElement('option');
        option.value = folder.id;
        option.textContent = folder.name;
        if (parentId === folder.id) {
            option.selected = true;
        }
        parentSelect.appendChild(option);
    });
    
    document.getElementById('folderModal').style.display = 'flex';
}

function saveFolder() {
    const name = document.getElementById('folderName').value.trim();
    const parentId = document.getElementById('parentFolder').value;
    
    if (!name) {
        alert('Vui lòng nhập tên folder');
        return;
    }
    
    // Check for duplicates
    const existing = folders.find(f => 
        f.name.toLowerCase() === name.toLowerCase() && 
        f.parentId === (parentId === 'INBOX' ? null : parentId)
    );
    
    if (existing) {
        alert('Đã có folder cùng tên trong thư mục này');
        return;
    }
    
    const newFolder = {
        id: generateId(),
        name: name,
        parentId: parentId === 'INBOX' ? null : parentId
    };
    
    folders.push(newFolder);
    closeFolderModal();
    renderFolders();
}

function closeFolderModal() {
    document.getElementById('folderModal').style.display = 'none';
}

function addChildFolder(parentId) {
    showAddFolderModal(parentId);
}

function deleteFolder(folderId) {
    if (!confirm('Bạn có chắc muốn xóa folder này và tất cả folder con?')) {
        return;
    }
    
    // Find all descendants
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
    
    // Remove rules that reference deleted folders
    rules = rules.filter(rule => {
        const folderActions = rule.actions?.filter(a => a.type === 'moveToFolder' || a.type === 'copyToFolder');
        if (folderActions && folderActions.some(a => toDelete.includes(a.value))) {
            return false;
        }
        return true;
    });
    
    renderFolders();
    renderRules();
}

function renderFolders() {
    const container = document.getElementById('foldersTree');
    
    if (folders.length === 0) {
        container.innerHTML = `
            <div class="empty-state" id="emptyFolders">
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
            const className = level === 0 ? '' : level === 1 ? 'child' : 'grandchild';
            
            html += `
                <div class="folder-item ${className}">
                    <div class="folder-icon">
                        <i class="fas fa-folder"></i>
                    </div>
                    <div class="folder-name">${folder.name}</div>
                    <div class="folder-type">${level === 0 ? 'Gốc' : level === 1 ? 'Cấp 1' : 'Cấp 2'}</div>
                    <div class="folder-actions">
                        <button class="action-btn add-child-btn" title="Thêm folder con" onclick="addChildFolder('${folder.id}')">
                            <i class="fas fa-plus"></i>
                        </button>
                        <button class="action-btn delete-btn" title="Xóa" onclick="deleteFolder('${folder.id}')">
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
    folders = [];
    rules = [];
    
    if (type === 'accounting') {
        // Accounting template
        const root1 = { id: generateId(), name: "01. HOA DON", parentId: null };
        folders.push(root1);
        
        folders.push({ id: generateId(), name: "DAU VAO", parentId: root1.id });
        folders.push({ id: generateId(), name: "DAU RA", parentId: root1.id });
        
        const root2 = { id: generateId(), name: "02. BAO CAO", parentId: null };
        folders.push(root2);
        
        folders.push({ id: generateId(), name: "THUE", parentId: root2.id });
        
        // Add accounting rules
        addAccountingRules();
        
    } else if (type === 'project') {
        // Project template
        const duAn = { id: generateId(), name: "DU AN A", parentId: null };
        folders.push(duAn);
        
        const tailieu = { id: generateId(), name: "01. TAILIEU", parentId: duAn.id };
        folders.push(tailieu);
        
        folders.push({ id: generateId(), name: "HOP DONG", parentId: tailieu.id });
        
        const lienhe = { id: generateId(), name: "02. LIEN HE", parentId: duAn.id };
        folders.push(lienhe);
        
        folders.push({ id: generateId(), name: "KHACH HANG", parentId: lienhe.id });
        
        // Add project rules
        addProjectRules();
        
    } else if (type === 'simple') {
        // Simple template
        folders.push({ id: generateId(), name: "01. CAN XU LY", parentId: null });
        folders.push({ id: generateId(), name: "02. DA XU LY", parentId: null });
        folders.push({ id: generateId(), name: "03. LUU TRU", parentId: null });
        
        // Add simple rules
        addSimpleRules();
    }
    
    renderFolders();
    renderRules();
}

function addAccountingRules() {
    const invoiceFolder = folders.find(f => f.name === "DAU VAO");
    const taxFolder = folders.find(f => f.name === "THUE");
    
    if (invoiceFolder) {
        rules.push({
            id: generateId(),
            name: "Hóa đơn đầu vào",
            conditions: [
                { type: 'subjectContainsWords', value: 'hóa đơn' },
                { type: 'hasAttachment', value: 'true' }
            ],
            actions: [
                { type: 'moveToFolder', value: invoiceFolder.id },
                { type: 'applyCategory', value: 'Hóa đơn' }
            ],
            enabled: true,
            expanded: false
        });
    }
    
    if (taxFolder) {
        rules.push({
            id: generateId(),
            name: "Email thuế",
            conditions: [
                { type: 'subjectContainsWords', value: 'thuế' },
                { type: 'importance', value: 'High' }
            ],
            actions: [
                { type: 'moveToFolder', value: taxFolder.id }
            ],
            enabled: true,
            expanded: false
        });
    }
}

function addKeyword(ruleId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    if (!rule.keywords) rule.keywords = [];
    
    const trimmedKeyword = keyword.trim();
    if (!trimmedKeyword) return;
    
    // Kiểm tra trùng lặp
    if (!rule.keywords.includes(trimmedKeyword)) {
        rule.keywords.push(trimmedKeyword);
        renderRules();
    }
}

function removeKeyword(ruleId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.keywords) {
        rule.keywords = rule.keywords.filter(k => k !== keyword);
        renderRules();
    }
}

function handleKeywordInputKeydown(event, ruleId) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const input = event.target;
        const keyword = input.value.trim();
        
        if (keyword) {
            addKeyword(ruleId, keyword);
            input.value = '';
        }
    } else if (event.key === 'Backspace' && event.target.value === '') {
        // Xóa tag cuối cùng khi nhấn Backspace với ô input trống
        const rule = rules.find(r => r.id === ruleId);
        if (rule && rule.keywords && rule.keywords.length > 0) {
            rule.keywords.pop();
            renderRules();
        }
    }
}

function handleKeywordInputBlur(event, ruleId) {
    const input = event.target;
    const keyword = input.value.trim();
    
    if (keyword) {
        addKeyword(ruleId, keyword);
        input.value = '';
    }
}

function renderKeywordsInput(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    const keywords = rule?.keywords || [];
    
    return `
        <div class="keywords-container">
            <label style="display: block; margin-bottom: 8px; font-weight: 500; color: #495057;">
                <i class="fas fa-tags"></i> Từ khóa (nhấn Enter để thêm)
            </label>
            <div class="keywords-input-container">
                ${keywords.map(keyword => `
                    <div class="keyword-tag">
                        ${keyword}
                        <button type="button" class="keyword-tag-remove" 
                                onclick="removeKeyword('${ruleId}', '${escapeHtml(keyword)}')">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                `).join('')}
                <input type="text" 
                       class="keywords-input" 
                       placeholder="Nhập từ khóa và nhấn Enter..."
                       onkeydown="handleKeywordInputKeydown(event, '${ruleId}')"
                       onblur="handleKeywordInputBlur(event, '${ruleId}')"
                       data-rule="${ruleId}">
            </div>
            <div class="keywords-instruction">
                <i class="fas fa-info-circle"></i>
                Các từ khóa này sẽ được sử dụng cho điều kiện "Subject includes" hoặc "Body includes"
            </div>
        </div>
    `;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function addProjectRules() {
    const contractFolder = folders.find(f => f.name === "HOP DONG");
    const customerFolder = folders.find(f => f.name === "KHACH HANG");
    
    if (contractFolder) {
        rules.push({
            id: generateId(),
            name: "Hợp đồng dự án",
            conditions: [
                { type: 'subjectContainsWords', value: 'hợp đồng' },
                { type: 'hasAttachment', value: 'true' }
            ],
            actions: [
                { type: 'moveToFolder', value: contractFolder.id },
                { type: 'applyCategory', value: 'Dự án' }
            ],
            enabled: true,
            expanded: false
        });
    }
    
    if (customerFolder) {
        rules.push({
            id: generateId(),
            name: "Email khách hàng",
            conditions: [
                { type: 'from', value: 'customer@' },
                { type: 'importance', value: 'High' }
            ],
            actions: [
                { type: 'moveToFolder', value: customerFolder.id }
            ],
            enabled: true,
            expanded: false
        });
    }
}

function addSimpleRules() {
    const processFolder = folders.find(f => f.name === "01. CAN XU LY");
    const doneFolder = folders.find(f => f.name === "02. DA XU LY");
    
    if (processFolder) {
        rules.push({
            id: generateId(),
            name: "Email khẩn cấp",
            conditions: [
                { type: 'importance', value: 'High' },
                { type: 'subjectContainsWords', value: 'urgent' }
            ],
            actions: [
                { type: 'moveToFolder', value: processFolder.id },
                { type: 'applyCategory', value: 'Khẩn cấp' }
            ],
            enabled: true,
            expanded: false
        });
    }
    
    if (doneFolder) {
        rules.push({
            id: generateId(),
            name: "Đã xử lý",
            conditions: [
                { type: 'subjectContainsWords', value: '[DONE]' }
            ],
            actions: [
                { type: 'moveToFolder', value: doneFolder.id },
                { type: 'markAsRead', value: '' }
            ],
            enabled: true,
            expanded: false
        });
    }
}

function clearFolders() {
    if (folders.length > 0 && !confirm('Bạn có chắc muốn xóa tất cả folder?')) {
        return;
    }
    folders = [];
    rules = [];
    renderFolders();
    renderRules();
}

// ==================== RULE FUNCTIONS ====================
function addNewRule() {
    const newRule = {
        id: generateId(),
        name: `Rule ${rules.length + 1}`,
        conditions: [],
        actions: [],
        keywords: [],  // THÊM MẢNG KEYWORDS
        enabled: true,
        expanded: false
    };
    rules.push(newRule);
    renderRules();
}

function deleteRule(ruleId) {
    if (!confirm('Bạn có chắc muốn xóa rule này?')) {
        return;
    }
    rules = rules.filter(r => r.id !== ruleId);
    renderRules();
}

function toggleRuleExpanded(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.expanded = !rule.expanded;
        renderRules();
    }
}

function addCondition(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    if (!rule.conditions) rule.conditions = [];
    
    rule.conditions.push({
        id: generateId(), // THÊM ID CHO MỖI CONDITION
        type: allConditions[0].id,
        value: '',
        keywords: [] // THÊM MẢNG KEYWORDS VÀO CONDITION
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


function updateCondition(ruleId, conditionId, type, value) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.conditions) {
        const condition = rule.conditions.find(c => c.id === conditionId);
        if (condition) {
            condition.type = type;
            condition.value = value;
        }
    }
}

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

function updateAction(ruleId, index, type, value) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule && rule.actions && rule.actions[index]) {
        rule.actions[index].type = type;
        rule.actions[index].value = value;
    }
}

function updateRuleName(ruleId, name) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.name = name;
    }
}

function toggleRuleEnabled(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.enabled = !rule.enabled;
        renderRules();
    }
}

function renderRules() {
    const container = document.getElementById('rulesContainer');
    
    if (rules.length === 0) {
        container.innerHTML = `
            <div class="empty-state" id="emptyRules">
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
        // Render conditions summary
        const conditionsText = rule.conditions ? rule.conditions.map(cond => {
            const condition = allConditions.find(c => c.id === cond.type);
            if (!condition) return '';
            
            let displayValue = cond.value;
            if (condition.param === 'boolean') {
                displayValue = cond.value === 'true' ? 'Có' : 'Không';
            } else if (condition.param === 'folder') {
                const folder = folders.find(f => f.id === cond.value);
                displayValue = folder ? folder.name : '(chọn folder)';
            }
            return `${condition.name}: ${displayValue || '(trống)'}`;
        }).filter(text => text).join(', ') : '';
        
        // Render actions summary
        const actionsText = rule.actions ? rule.actions.map(action => {
            const actionDef = allActions.find(a => a.id === action.type);
            if (!actionDef) return '';
            
            let displayValue = action.value;
            if (actionDef.param === 'folder') {
                const folder = folders.find(f => f.id === action.value);
                displayValue = folder ? folder.name : '(chọn folder)';
            }
            return `${actionDef.name}${displayValue ? `: ${displayValue}` : ''}`;
        }).filter(text => text).join(', ') : '';
        
        html += `
            <div class="rule-item">
                <div class="rule-header">
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <input type="checkbox" ${rule.enabled ? 'checked' : ''} onchange="toggleRuleEnabled('${rule.id}')">
                        <input type="text" value="${rule.name}" onchange="updateRuleName('${rule.id}', this.value)" 
                               style="border: 1px solid #ced4da; border-radius: 4px; padding: 4px 8px; font-size: 13px; width: 250px;">
                        <button class="toggle-rule-btn" onclick="toggleRuleExpanded('${rule.id}')">
                            <i class="fas fa-${rule.expanded ? 'chevron-up' : 'chevron-down'}"></i> ${rule.expanded ? 'Thu gọn' : 'Chi tiết'}
                        </button>
                    </div>
                    <button class="action-btn delete-btn" onclick="deleteRule('${rule.id}')">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
                
                <div class="conditions-summary">
                    <strong>Điều kiện:</strong> ${conditionsText || '<span style="color: #6c757d; font-style: italic;">Chưa có điều kiện</span>'}
                </div>
                
                <div class="actions-summary">
                    <strong>Hành động:</strong> ${actionsText || '<span style="color: #6c757d; font-style: italic;">Chưa có hành động</span>'}
                </div>
                
                ${rule.expanded ? renderRuleExpanded(rule) : ''}
            </div>
        `;
    });
    
    container.innerHTML = html;
}

function renderRuleExpanded(rule) {
    return `
        <div class="rule-expanded">
           
            
            <div class="condition-group" style="margin-bottom: 20px;">
                <div class="group-title">Điều kiện</div>
                ${rule.conditions ? rule.conditions.map((condition, index) => renderCondition(rule.id, index, condition)).join('') : 
                    '<div style="color: #6c757d; font-style: italic; padding: 10px;">Chưa có điều kiện nào</div>'}
                <button class="add-btn" onclick="addCondition('${rule.id}')">
                    <i class="fas fa-plus"></i> Thêm điều kiện
                </button>
            </div>
            
            <div class="action-group">
                <div class="group-title">Hành động</div>
                ${rule.actions ? rule.actions.map((action, index) => renderAction(rule.id, index, action)).join('') : 
                    '<div style="color: #6c757d; font-style: italic; padding: 10px;">Chưa có hành động nào</div>'}
                <button class="add-btn" onclick="addAction('${rule.id}')">
                    <i class="fas fa-plus"></i> Thêm hành động
                </button>
            </div>
        </div>
    `;
}

function addQuickKeywords(ruleId) {
    const quickKeywords = ['urgent', 'important', 'invoice', 'report', 'meeting', 'deadline'];
    
    return `
        <div style="margin-top: 10px;">
            <div style="font-size: 12px; color: #6c757d; margin-bottom: 5px;">Gợi ý:</div>
            <div class="quick-keywords">
                ${quickKeywords.map(keyword => `
                    <button type="button" class="quick-keyword-btn" 
                            onclick="addKeyword('${ruleId}', '${keyword}')">
                        ${keyword}
                    </button>
                `).join('')}
            </div>
        </div>
    `;
}

function renderCondition(ruleId, index, condition) {
    const conditionDef = allConditions.find(c => c.id === condition.type) || allConditions[0];
    
    let inputHtml = renderConditionInput(ruleId, condition.id, condition, conditionDef);
    
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
    // Các condition có param là 'text' sẽ hỗ trợ tags
    const supportsTags = conditionDef.param === 'text';
    
    if (supportsTags && !condition.keywords) {
        condition.keywords = [];
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
                       placeholder="${conditionDef.placeholder}"
                       style="width: 120px;">
                <span>KB</span>
            `;
        case 'date':
            return `
                <input type="text" class="condition-value-input" value="${condition.value || ''}" 
                       onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)"
                       placeholder="${conditionDef.placeholder}"
                       style="width: 150px;">
            `;
        case 'email':
            return `
                <input type="email" class="condition-value-input" value="${condition.value || ''}" 
                       onchange="updateCondition('${ruleId}', '${conditionId}', '${condition.type}', this.value)"
                       placeholder="${conditionDef.placeholder}">
            `;
        case 'none':
            return `<div class="condition-value-input" style="color: #6c757d; font-style: italic;">Không cần giá trị</div>`;
        case 'text':
            // ĐÂY LÀ PHẦN QUAN TRỌNG: RENDER TAG INPUT CHO TEXT CONDITIONS
            return renderConditionTagInput(ruleId, conditionId, condition, conditionDef);
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
            <div class="keywords-input-container" style="border: 1px solid #ced4da; border-radius: 4px; padding: 6px;">
                ${keywords.map(keyword => `
                    <div class="keyword-tag" style="margin: 2px;">
                        ${keyword}
                        <button type="button" class="keyword-tag-remove" 
                                onclick="removeConditionKeyword('${ruleId}', '${conditionId}', '${escapeHtml(keyword)}')">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                `).join('')}
                <input type="text" 
                       class="keywords-input" 
                       style="border: none; padding: 4px 6px; min-width: 100px;"
                       placeholder="${conditionDef.placeholder}..."
                       onkeydown="handleConditionKeywordKeydown(event, '${ruleId}', '${conditionId}')"
                       onblur="handleConditionKeywordBlur(event, '${ruleId}', '${conditionId}')"
                       data-rule="${ruleId}"
                       data-condition="${conditionId}">
            </div>
        </div>
    `;
}

// Thêm keyword vào condition
function addConditionKeyword(ruleId, conditionId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    const condition = rule.conditions?.find(c => c.id === conditionId);
    if (!condition) return;
    
    if (!condition.keywords) condition.keywords = [];
    
    const trimmedKeyword = keyword.trim();
    if (!trimmedKeyword) return;
    
    // Kiểm tra trùng lặp
    if (!condition.keywords.includes(trimmedKeyword)) {
        condition.keywords.push(trimmedKeyword);
        renderRules();
    }
}

// Xóa keyword khỏi condition
function removeConditionKeyword(ruleId, conditionId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (!rule) return;
    
    const condition = rule.conditions?.find(c => c.id === conditionId);
    if (condition && condition.keywords) {
        condition.keywords = condition.keywords.filter(k => k !== keyword);
        renderRules();
    }
}

// Xử lý keydown trong input
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
        // Xóa tag cuối cùng khi nhấn Backspace với ô input trống
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

function renderAction(ruleId, index, action) {
    const actionDef = allActions.find(a => a.id === action.type) || allActions[0];
    
    let inputHtml = renderActionInput(ruleId, index, action, actionDef);
    
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
            return `
                <input type="email" class="action-value-input" value="${action.value || ''}" 
                       onchange="updateAction('${ruleId}', ${index}, '${action.type}', this.value)"
                       placeholder="${actionDef.placeholder}">
            `;
        case 'text':
            return `
                <input type="text" class="action-value-input" value="${action.value || ''}" 
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

// ==================== SCRIPT GENERATION ====================
// ==================== SCRIPT GENERATION ====================
function generateScriptContent(userEmail, scriptName, additionalNotes) {
    const createAllFolders = document.getElementById('createAllFolders').checked;
    const enableAllRules = document.getElementById('enableAllRules').checked;
    
    if (!scriptName.toLowerCase().endsWith('.ps1')) {
        scriptName += '.ps1';
    }
    
    let script = `# ${scriptName}\n`;
    script += `# Outlook Rules Script - Tạo tự động\n`;
    script += `# Email: ${userEmail}\n`;
    script += `# Ngày tạo: ${new Date().toLocaleDateString('vi-VN')}\n`;
    
    if (additionalNotes) {
        script += `# Ghi chú: ${additionalNotes}\n`;
    }
    
    script += `\n# ==================== CẤU HÌNH ====================\n`;
    script += `# Cấu hình UTF-8 cho tiếng Việt\n`;
    script += `try {\n`;
    script += `    chcp 65001 | Out-Null\n`;
    script += `    \$OutputEncoding = [System.Text.Encoding]::UTF8\n`;
    script += `} catch { }\n\n`;
    
    script += `Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan\n`;
    script += `Write-Host "║         TẠO OUTLOOK RULES TỰ ĐỘNG                  ║" -ForegroundColor Cyan\n`;
    script += `Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan\n`;
    script += `Write-Host "Email: ${userEmail}" -ForegroundColor Yellow\n`;
    script += `Write-Host ""\n\n`;
    
    script += `# Kết nối Exchange Online\n`;
    script += `try {\n`;
    script += `    Connect-ExchangeOnline -UserPrincipalName "${userEmail}" -ShowProgress \$true\n`;
    script += `    Write-Host "✓ Đã kết nối Exchange Online" -ForegroundColor Green\n`;
    script += `} catch {\n`;
    script += `    Write-Host "✗ Lỗi kết nối Exchange Online" -ForegroundColor Red\n`;
    script += `    exit 1\n`;
    script += `}\n`;
    script += `\$mailbox = "${userEmail}"\n\n`;
    
    // Tạo folders
    if (folders.length > 0 && createAllFolders) {
        script += `# ==================== TẠO FOLDERS ====================\n`;
        script += `Write-Host "Tạo ${folders.length} folder..." -ForegroundColor Cyan\n`;
        
        // Dictionary lưu đường dẫn folder
        script += `\$folderPaths = @{}\n`;
        
        // Root folders first
        folders.filter(f => !f.parentId).forEach(folder => {
            const escapedName = escapePS(folder.name);
            script += `try {\n`;
            script += `    New-MailboxFolder -Name "${escapedName}" -Parent "\$mailbox\`:\\Inbox" -ErrorAction SilentlyContinue\n`;
            script += `    \$folderPaths['${folder.id}'] = "\$mailbox\`:\\Inbox\\${escapedName}"\n`;
            script += `    Write-Host "  ✓ ${folder.name}" -ForegroundColor Green\n`;
            script += `} catch {\n`;
            script += `    \$folderPaths['${folder.id}'] = "\$mailbox\`:\\Inbox\\${escapedName}"\n`;
            script += `    Write-Host "  ⚠ ${folder.name} (có thể đã tồn tại)" -ForegroundColor Yellow\n`;
            script += `}\n`;
        });
        
        // Child folders
        folders.filter(f => f.parentId).forEach(folder => {
            const parent = folders.find(p => p.id === folder.parentId);
            if (parent) {
                const escapedName = escapePS(folder.name);
                script += `try {\n`;
                script += `    \$parentPath = \$folderPaths['${parent.id}']\n`;
                script += `    if (\$parentPath) {\n`;
                script += `        New-MailboxFolder -Name "${escapedName}" -Parent "\$parentPath" -ErrorAction SilentlyContinue\n`;
                script += `        \$folderPaths['${folder.id}'] = "\$parentPath\\${escapedName}"\n`;
                script += `        Write-Host "  ✓ ${parent.name} → ${folder.name}" -ForegroundColor Green\n`;
                script += `    }\n`;
                script += `} catch {\n`;
                script += `    Write-Host "  ⚠ ${parent.name} → ${folder.name} (có thể đã tồn tại)" -ForegroundColor Yellow\n`;
                script += `}\n`;
            }
        });
        
        script += `Write-Host ""\n`;
    }
    
    // Tạo rules
    const enabledRules = rules.filter(r => r.enabled && enableAllRules);
    if (enabledRules.length > 0) {
        script += `# ==================== TẠO RULES ====================\n`;
        script += `Write-Host "Tạo ${enabledRules.length} rules..." -ForegroundColor Cyan\n`;
        
        enabledRules.forEach((rule, index) => {
            script += `\n# Rule ${index + 1}: ${rule.name}\n`;
            script += `try {\n`;
            script += `    Write-Host "Rule ${index + 1}: ${rule.name}" -NoNewline\n`;
            
            // Build New-InboxRule command
            let cmdParams = [];
            cmdParams.push(`-Name "${escapePS(rule.name)}"`);
            
            // Add conditions
            if (rule.conditions && rule.conditions.length > 0) {
    rule.conditions.forEach(condition => {
        const conditionDef = allConditions.find(c => c.id === condition.type);
        
        if (conditionDef) {
            // Xử lý các condition có keywords
            if (condition.keywords && condition.keywords.length > 0) {
                const keywordsString = condition.keywords.map(k => escapePS(k)).join(', ');
                
                switch(condition.type) {
                    case 'subjectContainsWords':
                        cmdParams.push(`-SubjectContainsWords "${keywordsString}"`);
                        break;
                    case 'subjectOrBodyContainsWords':
                        cmdParams.push(`-SubjectOrBodyContainsWords "${keywordsString}"`);
                        break;
                    case 'bodyContainsWords':
                        cmdParams.push(`-BodyContainsWords "${keywordsString}"`);
                        break;
                    case 'fromAddressContainsWords':
                        cmdParams.push(`-FromAddressContainsWords "${keywordsString}"`);
                        break;
                    // Các trường hợp khác vẫn dùng condition.value
                }
            } 
            // Xử lý các condition không có keywords
            else if (condition.value && condition.value.trim() !== '') {
                // ... phần xử lý condition.value như cũ
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
                                    // THÊM DẤU NGOẶC KÉP
                                    cmdParams.push(`-CopyToFolder "\$mailbox\`:\\Inbox\\${escapePS(folderPath)}"`);
                                }
                                break;
                            case 'deleteMessage':
                                cmdParams.push(`-DeleteMessage`);
                                break;
                            case 'markAsRead':
                                cmdParams.push(`-MarkAsRead`);
                                break;
                            case 'markAsJunk':
                                cmdParams.push(`-MarkAsJunk`);
                                break;
                            case 'setImportance':
                                cmdParams.push(`-SetImportance "${escapePS(action.value)}"`);
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
            
            // Không cần thêm -Enabled vì mặc định rule sẽ được enable
// Chỉ thêm -Enabled $false nếu rule bị disabled
if (!rule.enabled) {
    cmdParams.push(`-Enabled \$false`);
}
// Nếu rule.enabled = true thì không cần thêm gì cả
            
            // Tạo lệnh hoàn chỉnh
            const cmd = `    New-InboxRule ${cmdParams.join(' ')}`;
            
            script += cmd + `\n`;
            script += `    Write-Host " ✓" -ForegroundColor Green\n`;
            script += `} catch {\n`;
            script += `    Write-Host " ✗" -ForegroundColor Red\n`;
            script += `    Write-Host "    Lỗi: \$_" -ForegroundColor Red\n`;
            script += `    Write-Host "    Command đã thử: New-InboxRule ${cmdParams.slice(0, 3).join(' ')}..." -ForegroundColor Gray\n`;
            script += `}\n`;
        });
    }
    
    script += `\n# ==================== HOÀN TẤT ====================\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Green\n`;
    script += `Write-Host "║                ĐÃ HOÀN THÀNH!                      ║" -ForegroundColor Green\n`;
    script += `Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Green\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "Tổng số folder: ${folders.length}" -ForegroundColor Cyan\n`;
    script += `Write-Host "Tổng số rules: ${enabledRules.length}" -ForegroundColor Cyan\n`;
    script += `Write-Host ""\n`;
    
    script += `# Ngắt kết nối\n`;
    script += `Disconnect-ExchangeOnline -Confirm:\$false\n`;
    script += `Write-Host "Đã ngắt kết nối Exchange Online" -ForegroundColor Gray\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "Nhấn Enter để thoát..." -ForegroundColor Gray\n`;
    script += `pause\n`;
    
    return { script, fileName: scriptName };
}

function escapePS(text) {
    return text ? text.replace(/"/g, '`"').replace(/\$/g, '`$') : '';
}

// ==================== PREVIEW & DOWNLOAD ====================
function previewScript() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_outlook_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        alert('Vui lòng nhập email tài khoản Exchange Online');
        return;
    }
    
    const { script } = generateScriptContent(userEmail, scriptName, additionalNotes);
    document.getElementById('previewText').textContent = script;
    document.getElementById('previewModal').style.display = 'flex';
}

function closePreviewModal() {
    document.getElementById('previewModal').style.display = 'none';
}

function copyScriptToClipboard() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_outlook_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        alert('Vui lòng nhập email tài khoản Exchange Online');
        return;
    }
    
    const { script, fileName } = generateScriptContent(userEmail, scriptName, additionalNotes);
    
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(script).then(() => {
            alert(`✅ Đã sao chép script vào clipboard!\n\nLưu thành file: ${fileName}`);
        }).catch(() => {
            alert('❌ Lỗi sao chép, vui lòng thử lại');
        });
    } else {
        alert('Trình duyệt không hỗ trợ sao chép tự động, vui lòng dùng chức năng xem trước');
    }
}

function downloadScriptFile() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_outlook_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        alert('Vui lòng nhập email tài khoản Exchange Online');
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
    
    try {
        a.click();
        setTimeout(() => {
            if (document.body.contains(a)) {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        }, 1000);
        
        setTimeout(() => {
            alert(`✅ Đã tải xuống file: ${fileName}\n\nChạy script trong PowerShell với quyền Admin.`);
        }, 300);
        
    } catch (error) {
        console.error('Lỗi tải file:', error);
        document.getElementById('previewText').textContent = script;
        document.getElementById('previewModal').style.display = 'flex';
        alert('❌ Không thể tải file tự động, vui lòng sao chép từ cửa sổ xem trước.');
    }
}