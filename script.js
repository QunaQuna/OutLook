// ==================== GLOBAL VARIABLES ====================
let folders = [];
let rules = [];
let editingFolderId = null;
let isVSCode = false;

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', function() {
    console.log('Exchange Online Folder Generator loaded');
    
    // Kiểm tra môi trường
    isVSCode = window.location.protocol === 'vscode-webview:' || 
               window.location.href.includes('vscode-webview');
    
    if (isVSCode) {
        // Hiển thị cảnh báo cho VS Code
        document.getElementById('vscodeWarning').style.display = 'block';
        
        // Thay đổi text nút download
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.innerHTML = '<i class="fas fa-download"></i> Thử Tải File';
        downloadBtn.title = 'Trong VS Code, dùng nút "Sao chép Script" để đảm bảo';
        
        // Cập nhật hướng dẫn
        document.getElementById('instructionsList').innerHTML = `
            <li>Thiết lập folder và rules ở các mục bên dưới</li>
            <li><strong>Bạn đang dùng VS Code</strong> - tính năng tải file có thể không hoạt động</li>
            <li>Dùng nút "Sao chép Script" và tạo file thủ công</li>
            <li>Hoặc mở file này bằng trình duyệt (Chrome/Edge/Firefox)</li>
            <li>Chạy lệnh: <code>Unblock-File -Path "đường-dẫn-đến-file.ps1"</code></li>
            <li><strong>QUAN TRỌNG:</strong> Để chạy tiếng Việt, dùng Windows Terminal hoặc cấu hình UTF-8</li>
        `;
    }
    
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
    rules = rules.filter(r => !toDelete.includes(r.targetFolder));
    
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
    
    if (type === 'accounting') {
        // Accounting template
        const root1 = { id: generateId(), name: "01. HOA DON", parentId: null };
        folders.push(root1);
        
        folders.push({ id: generateId(), name: "DAU VAO", parentId: root1.id });
        folders.push({ id: generateId(), name: "DAU RA", parentId: root1.id });
        
        const root2 = { id: generateId(), name: "02. BAO CAO", parentId: null };
        folders.push(root2);
        
        folders.push({ id: generateId(), name: "THUE", parentId: root2.id });
        folders.push({ id: generateId(), name: "TAI CHINH", parentId: root2.id });
        
        const root3 = { id: generateId(), name: "03. HOP DONG", parentId: null };
        folders.push(root3);
        
        const khachHang = { id: generateId(), name: "KHACH HANG", parentId: root3.id };
        folders.push(khachHang);
        
        folders.push({ id: generateId(), name: "MOI", parentId: khachHang.id });
        folders.push({ id: generateId(), name: "DA KY", parentId: khachHang.id });
        folders.push({ id: generateId(), name: "NHA CUNG CAP", parentId: root3.id });
        
    } else if (type === 'project') {
        // Project template
        const duAn = { id: generateId(), name: "DU AN A", parentId: null };
        folders.push(duAn);
        
        const tailieu = { id: generateId(), name: "01. TAILIEU", parentId: duAn.id };
        folders.push(tailieu);
        
        folders.push({ id: generateId(), name: "HOP DONG", parentId: tailieu.id });
        folders.push({ id: generateId(), name: "BAO GIA", parentId: tailieu.id });
        
        const lienhe = { id: generateId(), name: "02. LIEN HE", parentId: duAn.id };
        folders.push(lienhe);
        
        folders.push({ id: generateId(), name: "KHACH HANG", parentId: lienhe.id });
        folders.push({ id: generateId(), name: "CHU DAU TU", parentId: lienhe.id });
        
    } else if (type === 'simple') {
        // Simple template
        folders.push({ id: generateId(), name: "01. CAN XU LY", parentId: null });
        folders.push({ id: generateId(), name: "02. DA XU LY", parentId: null });
        folders.push({ id: generateId(), name: "03. LUU TRU", parentId: null });
    }
    
    renderFolders();
    renderRules();
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
        keywords: [],
        targetFolder: folders.length > 0 ? folders[0].id : '',
        enabled: true
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

function addKeyword(ruleId) {
    const input = document.getElementById('keywordInput-' + ruleId);
    if (!input) return;
    
    const keyword = input.value.trim();
    if (!keyword) return;
    
    const rule = rules.find(r => r.id === ruleId);
    if (rule && !rule.keywords.includes(keyword)) {
        rule.keywords.push(keyword);
        input.value = '';
        renderRules();
    }
}

function removeKeyword(ruleId, keyword) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.keywords = rule.keywords.filter(k => k !== keyword);
        renderRules();
    }
}

function updateRuleName(ruleId, name) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.name = name;
    }
}

function updateRuleTarget(ruleId, folderId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.targetFolder = folderId;
    }
}

function toggleRuleEnabled(ruleId) {
    const rule = rules.find(r => r.id === ruleId);
    if (rule) {
        rule.enabled = !rule.enabled;
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
        const targetFolder = folders.find(f => f.id === rule.targetFolder);
        
        html += `
            <div class="rule-item">
                <div class="rule-header">
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <input type="checkbox" ${rule.enabled ? 'checked' : ''} onchange="toggleRuleEnabled('${rule.id}')">
                        <input type="text" value="${rule.name}" onchange="updateRuleName('${rule.id}', this.value)" 
                               style="border: 1px solid #ced4da; border-radius: 4px; padding: 4px 8px; font-size: 13px; width: 200px;">
                    </div>
                    <button class="action-btn delete-btn" onclick="deleteRule('${rule.id}')">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
                
                <div class="form-group">
                    <label style="font-size: 13px;">Từ khóa (phân cách bằng dấu phẩy):</label>
                    <div id="keywords-${rule.id}" style="margin-bottom: 8px;">
                        ${rule.keywords.map(keyword => 
                            `<span class="keyword-tag">${keyword} 
                                <i class="fas fa-times" onclick="removeKeyword('${rule.id}', '${keyword}')"></i>
                            </span>`
                        ).join('')}
                        ${rule.keywords.length === 0 ? '<span style="color: #6c757d; font-style: italic; font-size: 12px;">Chưa có từ khóa</span>' : ''}
                    </div>
                    <div class="keyword-input-container">
                        <input type="text" id="keywordInput-${rule.id}" placeholder="Nhập từ khóa..." onkeypress="if(event.key==='Enter'){addKeyword('${rule.id}')}">
                        <button class="add-keyword-btn" onclick="addKeyword('${rule.id}')">
                            <i class="fas fa-plus"></i>
                        </button>
                    </div>
                </div>
                
                <div class="form-group">
                    <label style="font-size: 13px;">Folder đích:</label>
                    <select onchange="updateRuleTarget('${rule.id}', this.value)" style="font-size: 13px;">
                        <option value="">-- Chọn folder --</option>
                        ${folders.map(folder => 
                            `<option value="${folder.id}" ${rule.targetFolder === folder.id ? 'selected' : ''}>
                                ${getFolderPath(folder.id)}
                            </option>`
                        ).join('')}
                    </select>
                </div>
            </div>
        `;
    });
    
    container.innerHTML = html;
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

// ==================== SCRIPT GENERATION (ĐÃ SỬA ĐỂ HỖ TRỢ TIẾNG VIỆT) ====================
// ==================== SCRIPT GENERATION (ĐÃ TÍCH HỢP CẤU HÌNH UTF-8) ====================
function generateScriptContent(userEmail, scriptName, additionalNotes) {
    const createAllFolders = document.getElementById('createAllFolders').checked;
    const enableAllRules = document.getElementById('enableAllRules').checked;
    
    // Đảm bảo có đuôi .ps1
    if (!scriptName.toLowerCase().endsWith('.ps1')) {
        scriptName += '.ps1';
    }
    
    let script = `# ${scriptName}\n`;
    script += `# Script được tạo tự động từ Exchange Online Folder Rules Generator\n`;
    script += `# Email tài khoản: ${userEmail}\n`;
    script += `# Ngày tạo: ${new Date().toLocaleDateString('vi-VN')}\n`;
    
    if (additionalNotes) {
        script += `# Ghi chú: ${additionalNotes}\n`;
    }
    
    script += `\n# ==================== CẤU HÌNH TỰ ĐỘNG CHO TIẾNG VIỆT ====================\n`;
    script += `# Tự động cấu hình PowerShell để hiển thị tiếng Việt\n`;
    script += `try {\n`;
    script += `    # Kiểm tra và thay đổi code page sang UTF-8\n`;
    script += `    \$currentCodePage = chcp\n`;
    script += `    if (\$currentCodePage -ne 65001) {\n`;
    script += `        chcp 65001 | Out-Null\n`;
    script += `        Write-Host "[UTF-8] Đã đặt code page sang UTF-8" -ForegroundColor Green\n`;
    script += `    }\n`;
    script += `    \n`;
    script += `    # Cấu hình encoding cho PowerShell\n`;
    script += `    \$OutputEncoding = [System.Text.Encoding]::UTF8\n`;
    script += `    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n`;
    script += `    [Console]::InputEncoding = [System.Text.Encoding]::UTF8\n`;
    script += `    \n`;
    script += `    # Kiểm tra và đề xuất font nếu cần\n`;
    script += `    \$fontName = (Get-Host).UI.RawUI.Font.Name\n`;
    script += `    \$supportedFonts = @('Consolas', 'Cascadia Code', 'Lucida Console', 'DejaVu Sans Mono')\n`;
    script += `    \n`;
    script += `    if (\$supportedFonts -notcontains \$fontName) {\n`;
    script += `        Write-Host "[FONT] Đang dùng font: \$fontName" -ForegroundColor Yellow\n`;
    script += `        Write-Host "[FONT] Khuyến nghị dùng Consolas hoặc Cascadia Code để hiển thị tiếng Việt tốt nhất" -ForegroundColor Yellow\n`;
    script += `    } else {\n`;
    script += `        Write-Host "[FONT] Font hỗ trợ tiếng Việt: \$fontName" -ForegroundColor Green\n`;
    script += `    }\n`;
    script += `    \n`;
    script += `} catch {\n`;
    script += `    Write-Warning "Không thể cấu hình UTF-8 tự động: \$_"\n`;
    script += `    Write-Host "Vui lòng chạy lệnh thủ công: chcp 65001" -ForegroundColor Yellow\n`;
    script += `}\n\n`;
    
    script += `# ==================== THIẾT LẬP MÔI TRƯỜNG ====================\n`;
    script += `# Hiển thị thông tin bắt đầu\n`;
    script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Cyan\n`;
    script += `Write-Host "              TẠO FOLDER VÀ RULES TRÊN MAILBOX              " -ForegroundColor Cyan\n`;
    script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Cyan\n`;
    script += `Write-Host "Email: ${userEmail}" -ForegroundColor Yellow\n`;
    script += `Write-Host "Thời gian: \$(Get-Date -Format 'HH:mm:ss dd/MM/yyyy')" -ForegroundColor Gray\n`;
    script += `Write-Host ""\n\n`;
    
    script += `# Kiểm tra module Exchange Online\n`;
    script += `Write-Host "[1/3] Kiểm tra môi trường..." -ForegroundColor Cyan\n`;
    script += `try {\n`;
    script += `    \$exchangeModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable\n`;
    script += `    if (-not \$exchangeModule) {\n`;
    script += `        Write-Host "   ⚠ Module ExchangeOnlineManagement chưa được cài đặt" -ForegroundColor Yellow\n`;
    script += `        Write-Host "   Đang cài đặt module..." -ForegroundColor Yellow\n`;
    script += `        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser\n`;
    script += `        Write-Host "   ✓ Đã cài đặt module" -ForegroundColor Green\n`;
    script += `    } else {\n`;
    script += `        Write-Host "   ✓ Module ExchangeOnlineManagement đã sẵn sàng (v\$(\$exchangeModule.Version))" -ForegroundColor Green\n`;
    script += `    }\n`;
    script += `} catch {\n`;
    script += `    Write-Host "   ✗ Lỗi kiểm tra module: \$_" -ForegroundColor Red\n`;
    script += `    exit 1\n`;
    script += `}\n`;
    script += `Write-Host ""\n\n`;
    
    script += `# Kết nối Exchange Online\n`;
    script += `Write-Host "[2/3] Kết nối Exchange Online..." -ForegroundColor Cyan\n`;
    script += `try {\n`;
    script += `    Connect-ExchangeOnline -UserPrincipalName "${userEmail}" -ShowProgress \$true\n`;
    script += `    Write-Host "   ✓ Đã kết nối thành công" -ForegroundColor Green\n`;
    script += `} catch {\n`;
    script += `    Write-Host "   ✗ Không thể kết nối Exchange Online" -ForegroundColor Red\n`;
    script += `    Write-Host "   Chi tiết: \$_" -ForegroundColor Red\n`;
    script += `    exit 1\n`;
    script += `}\n`;
    script += `\$myMailbox = "${userEmail}"\n`;
    script += `Write-Host ""\n\n`;
    
    // Hàm helper để escape tên folder
    script += `# Hàm hỗ trợ xử lý tên folder\n`;
    script += `function Format-FolderName {\n`;
    script += `    param([string]\$Name)\n`;
    script += `    # Giữ nguyên tên folder, chỉ escape ký tự đặc biệt\n`;
    script += `    return \$Name.Replace('"', '\`"').Replace('\\', '\\\\')\n`;
    script += `}\n\n`;
    
    // Create folders
    if (folders.length > 0 && createAllFolders) {
        script += `Write-Host "[3/3] Đang tạo ${folders.length} folder..." -ForegroundColor Cyan\n`;
        script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor DarkGray\n\n`;
        
        // Tạo dictionary để lưu path của các folder
        script += `# Tạo dictionary lưu đường dẫn folder\n`;
        script += `\$folderPaths = @{}\n\n`;
        
        // Root folders first
        folders.filter(f => !f.parentId).forEach(folder => {
            const folderPath = getFolderPathForScript(folder.id);
            script += `# Tạo folder: ${folder.name}\n`;
            script += `try {\n`;
            script += `    Write-Host "• ${folder.name}" -NoNewline -ForegroundColor White\n`;
            script += `    \$folderName = Format-FolderName "${folder.name}"\n`;
            script += `    New-MailboxFolder -Name "\$folderName" -Parent "\$myMailbox\`:\\Inbox" -ErrorAction Stop\n`;
            script += `    \$folderPaths['${folder.id}'] = "\$myMailbox\`:\\Inbox\\\$folderName"\n`;
            script += `    Write-Host " ✓" -ForegroundColor Green\n`;
            script += `} catch {\n`;
            script += `    if (\$_.Exception.Message -like "*folder already exists*") {\n`;
            script += `        Write-Host " ⚠ (đã tồn tại)" -ForegroundColor Yellow\n`;
            script += `        \$folderPaths['${folder.id}'] = "\$myMailbox\`:\\Inbox\\\$(Format-FolderName '${folder.name}')"\n`;
            script += `    } else {\n`;
            script += `        Write-Host " ✗" -ForegroundColor Red\n`;
            script += `        Write-Host "   Lỗi: \$(\$_.Exception.Message)" -ForegroundColor Red\n`;
            script += `    }\n`;
            script += `}\n`;
        });
        
        script += `\n`;
        
        // Child folders
        const childFolders = folders.filter(f => f.parentId);
        if (childFolders.length > 0) {
            script += `# Tạo folder con\n`;
            
            childFolders.forEach(folder => {
                const parent = folders.find(p => p.id === folder.parentId);
                if (parent) {
                    const folderPath = getFolderPathForScript(folder.id);
                    const displayPath = getFolderPath(folder.id).replace(/→/g, '→');
                    
                    script += `# Tạo folder: ${displayPath}\n`;
                    script += `try {\n`;
                    script += `    Write-Host "• ${displayPath}" -NoNewline -ForegroundColor Gray\n`;
                    script += `    \$parentPath = \$folderPaths['${parent.id}']\n`;
                    script += `    if (\$parentPath) {\n`;
                    script += `        \$folderName = Format-FolderName "${folder.name}"\n`;
                    script += `        New-MailboxFolder -Name "\$folderName" -Parent "\$parentPath" -ErrorAction Stop\n`;
                    script += `        \$folderPaths['${folder.id}'] = "\$parentPath\\\$folderName"\n`;
                    script += `        Write-Host " ✓" -ForegroundColor Green\n`;
                    script += `    } else {\n`;
                    script += `        Write-Host " ✗ (không tìm thấy folder cha)" -ForegroundColor Red\n`;
                    script += `    }\n`;
                    script += `} catch {\n`;
                    script += `    if (\$_.Exception.Message -like "*folder already exists*") {\n`;
                    script += `        Write-Host " ⚠ (đã tồn tại)" -ForegroundColor Yellow\n`;
                    script += `        \$folderPaths['${folder.id}'] = "\$folderPaths['${parent.id}']\\\$(Format-FolderName '${folder.name}')"\n`;
                    script += `    } else {\n`;
                    script += `        Write-Host " ✗" -ForegroundColor Red\n`;
                    script += `        Write-Host "   Lỗi: \$(\$_.Exception.Message)" -ForegroundColor Red\n`;
                    script += `    }\n`;
                    script += `}\n`;
                }
            });
        }
        
        script += `\n`;
    }
    
    // Create rules
    const enabledRules = rules.filter(r => r.enabled && enableAllRules);
    if (enabledRules.length > 0) {
        script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor DarkGray\n`;
        script += `Write-Host "Đang tạo ${enabledRules.length} rules..." -ForegroundColor Cyan\n`;
        script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor DarkGray\n\n`;
        
        enabledRules.forEach((rule, index) => {
            if (rule.keywords.length > 0 && rule.targetFolder) {
                const targetFolder = folders.find(f => f.id === rule.targetFolder);
                if (targetFolder) {
                    const displayPath = getFolderPath(targetFolder.id).replace(/→/g, '→');
                    
                    script += `# Rule ${index + 1}: ${rule.name}\n`;
                    script += `try {\n`;
                    script += `    Write-Host "[Rule ${index + 1}] ${rule.name}" -NoNewline -ForegroundColor White\n`;
                    script += `    Write-Host " → ${displayPath}" -ForegroundColor Gray\n`;
                    
                    // Format keywords
                    const keywordsFormatted = rule.keywords.map(k => 
                        `"$(Format-FolderName "${k}")"`
                    ).join(', ');
                    
                    script += `    Write-Host "   Từ khóa: ${rule.keywords.join(', ')}" -ForegroundColor DarkGray\n`;
                    
                    // Get target folder path
                    script += `    \$targetPath = \$folderPaths['${targetFolder.id}']\n`;
                    script += `    if (\$targetPath) {\n`;
                    script += `        \$ruleName = Format-FolderName "${rule.name}"\n`;
                    script += `        New-InboxRule -Name "\$ruleName" \`\n`;
                    script += `            -SubjectContainsWords ${keywordsFormatted} \`\n`;
                    script += `            -MoveToFolder "\$targetPath"\n`;
                    script += `        Write-Host "   ✓ Đã tạo rule" -ForegroundColor Green\n`;
                    script += `    } else {\n`;
                    script += `        Write-Host "   ✗ Không tìm thấy folder đích" -ForegroundColor Red\n`;
                    script += `    }\n`;
                    script += `} catch {\n`;
                    script += `    Write-Host "   ✗ Lỗi tạo rule" -ForegroundColor Red\n`;
                    script += `    Write-Host "     Chi tiết: \$(\$_.Exception.Message)" -ForegroundColor Red\n`;
                    script += `}\n`;
                    script += `Write-Host ""\n`;
                }
            }
        });
    }
    
    script += `# ==================== KẾT THÚC ====================\n`;
    script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Green\n`;
    script += `Write-Host "                     HOÀN THÀNH!                           " -ForegroundColor Green\n`;
    script += `Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Green\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "✓ Đã hoàn thành tất cả thao tác" -ForegroundColor Green\n`;
    script += `Write-Host "✓ Tổng số folder đã xử lý: ${folders.length}" -ForegroundColor Cyan\n`;
    script += `Write-Host "✓ Tổng số rules đã tạo: ${enabledRules.length}" -ForegroundColor Cyan\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "Lưu ý:" -ForegroundColor Yellow\n`;
    script += `Write-Host "- Các folder và rules sẽ có hiệu lực ngay lập tức" -ForegroundColor Gray\n`;
    script += `Write-Host "- Kiểm tra lại trong Outlook hoặc OWA để xác nhận" -ForegroundColor Gray\n`;
    script += `Write-Host ""\n`;
    
    script += `# Ngắt kết nối\n`;
    script += `try {\n`;
    script += `    Disconnect-ExchangeOnline -Confirm:\$false -ErrorAction SilentlyContinue\n`;
    script += `    Write-Host "✓ Đã ngắt kết nối Exchange Online" -ForegroundColor Green\n`;
    script += `} catch {\n`;
    script += `    Write-Host "⚠ Không thể ngắt kết nối tự động" -ForegroundColor Yellow\n`;
    script += `}\n`;
    
    script += `\n# Giữ cửa sổ mở để xem kết quả\n`;
    script += `Write-Host ""\n`;
    script += `Write-Host "Nhấn Enter để thoát..." -ForegroundColor Gray\n`;
    script += `pause\n`;
    
    return { script, fileName: scriptName };
}

// Hàm helper để escape string cho PowerShell (giữ nguyên tiếng Việt)
function escapePowerShellString(text) {
    if (!text) return '';
    
    // Chỉ escape các ký tự đặc biệt, giữ nguyên tiếng Việt
    return text
        .replace(/\\/g, '\\\\')        // Escape backslashes
        .replace(/"/g, '`"')           // Escape double quotes
        .replace(/\$/g, '`$')          // Escape dollar sign
        .replace(/`/g, '``')           // Escape backtick
        .replace(/\r?\n/g, ' ');       // Replace newlines with space
}

// ==================== PREVIEW SCRIPT ====================
function previewScript() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_folders_rules';
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

// ==================== COPY TO CLIPBOARD ====================
function copyScriptToClipboard() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_folders_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        alert('Vui lòng nhập email tài khoản Exchange Online');
        return;
    }
    
    const { script, fileName } = generateScriptContent(userEmail, scriptName, additionalNotes);
    
    // Thông báo hướng dẫn về tiếng Việt
    let vietnameseGuide = `\n\n=== HƯỚNG DẪN CHẠY SCRIPT VỚI TIẾNG VIỆT ===\n`;
    vietnameseGuide += `1. Dùng Windows Terminal (khuyên dùng) hoặc PowerShell ISE\n`;
    vietnameseGuide += `2. Cấu hình font: Consolas hoặc Cascadia Code\n`;
    vietnameseGuide += `3. Chạy lệnh cấu hình trước: chcp 65001\n`;
    vietnameseGuide += `4. Đảm bảo file script được lưu với UTF-8 BOM\n`;
    
    const fullMessage = script + vietnameseGuide;
    
    // Try to use clipboard API
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(fullMessage).then(function() {
            let message = `✅ Đã sao chép script vào clipboard!\n\n`;
            message += `ĐỂ CHẠY VỚI TIẾNG VIỆT:\n`;
            message += `1. Mở PowerShell Admin\n`;
            message += `2. Cài đặt font Consolas hoặc Cascadia Code\n`;
            message += `3. Chạy: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser\n`;
            message += `4. Tạo file ${fileName} với encoding UTF-8\n`;
            message += `5. Chạy: Unblock-File -Path "đường-dẫn\\${fileName}"\n`;
            message += `6. Chạy script: .\\${fileName}\n`;
            message += `\nLƯU Ý: Dùng Windows Terminal để hiển thị tiếng Việt tốt nhất`;
            alert(message);
        }).catch(function(err) {
            fallbackCopy(script, fileName);
        });
    } else {
        fallbackCopy(script, fileName);
    }
}

function fallbackCopy(script, fileName) {
    // Fallback method for older browsers
    const textArea = document.createElement('textarea');
    textArea.value = script;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    textArea.style.top = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            let message = `✅ Đã sao chép script (phương pháp cũ)!\n\n`;
            message += `ĐỂ CHẠY VỚI TIẾNG VIỆT:\n`;
            message += `1. Mở Notepad++ hoặc VS Code\n`;
            message += `2. Dán script (Ctrl+V) và lưu với tên: ${fileName}\n`;
            message += `3. Chọn encoding: UTF-8 với BOM\n`;
            message += `4. Mở PowerShell Admin\n`;
            message += `5. Chạy: chcp 65001 (để đặt encoding UTF-8)\n`;
            message += `6. Chạy: .\\${fileName}`;
            alert(message);
        } else {
            document.getElementById('previewText').textContent = script;
            document.getElementById('previewModal').style.display = 'flex';
            alert(`⚠ Không thể sao chép tự động. Script đã được hiển thị trong cửa sổ xem trước. Hãy sao chép thủ công và lưu thành file ${fileName}`);
        }
    } catch (err) {
        console.error('Lỗi khi sao chép: ', err);
        document.getElementById('previewText').textContent = script;
        document.getElementById('previewModal').style.display = 'flex';
        alert(`❌ Lỗi khi sao chép. Script đã được hiển thị trong cửa sổ xem trước. Hãy sao chép thủ công và lưu thành file ${fileName}`);
    }
    
    document.body.removeChild(textArea);
}

// ==================== DOWNLOAD FILE (ĐÃ SỬA ĐỂ HỖ TRỢ UTF-8 BOM) ====================
function downloadScriptFile() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_folders_rules';
    const additionalNotes = document.getElementById('additionalNotes').value;
    
    if (!userEmail) {
        alert('Vui lòng nhập email tài khoản Exchange Online');
        return;
    }
    
    const { script, fileName } = generateScriptContent(userEmail, scriptName, additionalNotes);
    
    // Tạo blob từ script với UTF-8 BOM để hỗ trợ tiếng Việt
    const BOM = '\uFEFF'; // UTF-8 BOM
    const blob = new Blob([BOM + script], { 
        type: 'text/plain;charset=utf-8' 
    });
    const url = URL.createObjectURL(blob);
    
    // Tạo link tải xuống
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.style.display = 'none';
    
    // Thêm vào document
    document.body.appendChild(a);
    
    // Phương pháp đơn giản nhất
    try {
        a.click();
        
        // Dọn dẹp sau một khoảng thời gian
        setTimeout(() => {
            if (document.body.contains(a)) {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        }, 1000);
        
        // Hiển thị thông báo với hướng dẫn tiếng Việt
        let message = `✅ Đã khởi tạo tải xuống file: ${fileName}\n\n`;
        
        if (isVSCode) {
            message += `LƯU Ý: Bạn đang dùng VS Code. Nếu không thấy hộp thoại lưu file:\n`;
            message += `1. Dùng nút "Sao chép Script" ở trên\n`;
            message += `2. Hoặc mở file HTML này bằng Chrome/Edge/Firefox\n`;
            message += `3. Hoặc tạo file thủ công từ script đã sao chép\n\n`;
        } else {
            message += `BƯỚC TIẾP THEO:\n`;
            message += `1. Chọn nơi lưu file trong hộp thoại\n`;
        }
        
        message += `\n=== ĐỂ CHẠY VỚI TIẾNG VIỆT ===\n`;
        message += `2. Mở Windows Terminal hoặc PowerShell ISE\n`;
        message += `3. Cấu hình font: Consolas hoặc Cascadia Code\n`;
        message += `4. Chạy PowerShell với quyền Admin\n`;
        message += `5. Chạy lệnh: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser\n`;
        message += `6. Chạy: Unblock-File -Path "đường-dẫn\\${fileName}"\n`;
        message += `7. Chạy script: .\\${fileName}\n\n`;
        message += `📝 Lưu ý: Nếu không thấy tiếng Việt, hãy chạy lệnh 'chcp 65001' trước khi chạy script`;
        
        setTimeout(() => {
            alert(message);
        }, 300);
        
    } catch (error) {
        console.error('Lỗi khi tải file:', error);
        
        // Fallback: hiển thị script để copy thủ công
        document.getElementById('previewText').textContent = script;
        document.getElementById('previewModal').style.display = 'flex';
        
        alert(`❌ Không thể tải file tự động.\n\nScript đã được hiển thị trong cửa sổ xem trước.\nVui lòng sao chép và tạo file ${fileName} thủ công.\n\nLƯU Ý: Lưu file với encoding UTF-8 với BOM để hiển thị tiếng Việt.`);
        
        // Dọn dẹp
        if (document.body.contains(a)) {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
    }
}

// ==================== THÊM TÍNH NĂNG TẠO BATCH FILE ====================
function createBatchFileWrapper() {
    const userEmail = document.getElementById('userEmail').value;
    const scriptName = document.getElementById('scriptName').value || 'create_folders_rules';
    
    if (!userEmail) {
        alert('Vui lòng nhập email tài khoản Exchange Online');
        return;
    }
    
    // Tạo file batch để chạy script với UTF-8
    const batchContent = `@echo off
chcp 65001 > nul
echo ========================================
echo    CHAY SCRIPT EXCHANGE ONLINE
echo    (Hỗ trợ tiếng Việt - UTF-8)
echo ========================================
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0${scriptName}"
pause`;
    
    const batchBlob = new Blob(['\uFEFF' + batchContent], {
        type: 'text/plain;charset=utf-8'
    });
    const batchUrl = URL.createObjectURL(batchBlob);
    
    const batchLink = document.createElement('a');
    batchLink.href = batchUrl;
    batchLink.download = 'run_script.bat';
    batchLink.style.display = 'none';
    document.body.appendChild(batchLink);
    batchLink.click();
    
    setTimeout(() => {
        document.body.removeChild(batchLink);
        URL.revokeObjectURL(batchUrl);
    }, 1000);
    
    alert(`✅ Đã tạo file batch 'run_script.bat'\n\nChạy file này để tự động cấu hình UTF-8 và chạy script PowerShell.`);
}