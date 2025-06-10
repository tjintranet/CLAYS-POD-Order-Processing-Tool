function showInventoryTab() {
    // Check if inventory access is required and granted
    if (INVENTORY_ACCESS.enabled && !hasInventoryAccess()) {
        if (!requestInventoryAccess()) {
            return; // Stay on current tab if access denied
        }
    }
    
    const inventoryTab = document.getElementById('inventory-tab');
    if (inventoryTab && window.bootstrap && bootstrap.Tab) {
        const tab = new bootstrap.Tab(inventoryTab);
        tab.show();
        
        // Show access status bar if protection is enabled
        if (INVENTORY_ACCESS.enabled) {
            showAccessStatusBar();
            showInventoryStatus('✅ Inventory management access granted', 'success');
        }
    } else {
        // Fallback if Bootstrap JS isn't loaded
        const inventoryPane = document.getElementById('inventory-pane');
        const ordersPane = document.getElementById('orders-pane');
        if (inventoryPane && ordersPane) {
            ordersPane.classList.remove('show', 'active');
            inventoryPane.classList.add('show', 'active');
            inventoryTab.classList.add('active');
            document.getElementById('orders-tab').classList.remove('active');
            
            if (INVENTORY_ACCESS.enabled) {
                showAccessStatusBar();
            }
        }
    }
}

// Show the access status bar
function showAccessStatusBar() {
    const accessBar = document.getElementById('accessStatusBar');
    const accessText = document.getElementById('accessStatusText');
    
    if (accessBar && accessText) {
        try {
            const session = localStorage.getItem(INVENTORY_ACCESS.sessionKey);
            if (session) {
                const sessionData = JSON.parse(session);
                const expiresAt = new Date(sessionData.expires);
                const timeLeft = sessionData.expires - Date.now();
                const hoursLeft = Math.floor(timeLeft / (60 * 60 * 1000));
                const minutesLeft = Math.floor((timeLeft % (60 * 60 * 1000)) / (60 * 1000));
                
                accessText.textContent = ` - Session expires in ${hoursLeft}h ${minutesLeft}m`;
                accessBar.style.display = 'block';
            }
        } catch (error) {
            secureLog('Error showing access status:', error);
        }
    }
}

// Hide the access status bar  
function hideAccessStatusBar() {
    const accessBar = document.getElementById('accessStatusBar');
    if (accessBar) {
        accessBar.style.display = 'none';
    }
}

// Check if user has current inventory access
function hasInventoryAccess() {
    if (!INVENTORY_ACCESS.enabled) return true;
    
    try {
        const session = localStorage.getItem(INVENTORY_ACCESS.sessionKey);
        if (!session) return false;
        
        const sessionData = JSON.parse(session);
        const now = Date.now();
        return sessionData.expires > now;
    } catch (error) {
        secureLog('Session check error:', error);
        localStorage.removeItem(INVENTORY_ACCESS.sessionKey);
        return false;
    }
}

// Request inventory access with password
function requestInventoryAccess() {
    const password = prompt("🔐 Enter inventory management password:");
    if (!password) {
        showInventoryStatus('Access cancelled', 'info');
        return false;
    }
    
    return validatePasswordAndGrantAccess(password);
}

// Validate password and grant access
async function validatePasswordAndGrantAccess(password) {
    try {
        // Create SHA-256 hash of entered password
        const encoder = new TextEncoder();
        const data = encoder.encode(password);
        const hashBuffer = await crypto.subtle.digest('SHA-256', data);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        
        if (hashHex === INVENTORY_ACCESS.passwordHash) {
            // Grant access for session duration
            const sessionData = {
                granted: true,
                expires: Date.now() + INVENTORY_ACCESS.sessionDuration,
                grantedAt: new Date().toISOString()
            };
            localStorage.setItem(INVENTORY_ACCESS.sessionKey, JSON.stringify(sessionData));
            
            secureLog('Inventory access granted');
            return true;
        } else {
            alert("❌ Incorrect password. Access denied.");
            secureLog('Inventory access denied - incorrect password');
            return false;
        }
    } catch (error) {
        console.error('Password validation error:', error);
        alert("❌ Error validating password. Please try again.");
        return false;
    }
}

// Check if inventory functions should be protected
function requireInventoryAccess(functionName) {
    if (INVENTORY_ACCESS.enabled && !hasInventoryAccess()) {
        showInventoryStatus(`🔒 Access required for ${functionName}. Please authenticate first.`, 'warning');
        return false;
    }
    return true;
}

// Check session status and warn about expiration
function checkInventorySession() {
    if (!INVENTORY_ACCESS.enabled) return;
    
    try {
        const session = localStorage.getItem(INVENTORY_ACCESS.sessionKey);
        if (!session) {
            hideAccessStatusBar();
            return;
        }
        
        const sessionData = JSON.parse(session);
        const timeLeft = sessionData.expires - Date.now();
        const minutesLeft = Math.floor(timeLeft / (60 * 1000));
        
        // Update access status bar if visible
        const accessBar = document.getElementById('accessStatusBar');
        if (accessBar && accessBar.style.display !== 'none') {
            showAccessStatusBar();
        }
        
        // Show warning when 10 minutes or less remaining
        if (timeLeft <= 10 * 60 * 1000 && timeLeft > 0) {
            showInventoryStatus(`⏰ Inventory access expires in ${minutesLeft} minutes`, 'warning');
        } 
        // Session expired
        else if (timeLeft <= 0) {
            localStorage.removeItem(INVENTORY_ACCESS.sessionKey);
            hideAccessStatusBar();
            showInventoryStatus('🔒 Inventory access expired. Please re-authenticate.', 'danger');
            
            // Switch back to orders tab
            const ordersTab = document.getElementById('orders-tab');
            if (ordersTab && window.bootstrap && bootstrap.Tab) {
                const tab = new bootstrap.Tab(ordersTab);
                tab.show();
            }
        }
    } catch (error) {
        secureLog('Session check error:', error);
        localStorage.removeItem(INVENTORY_ACCESS.sessionKey);
        hideAccessStatusBar();
    }
}

// Extend session if user is active
function extendInventorySession() {
    if (!INVENTORY_ACCESS.enabled || !hasInventoryAccess()) return;
    
    try {
        const session = localStorage.getItem(INVENTORY_ACCESS.sessionKey);
        if (session) {
            const sessionData = JSON.parse(session);
            sessionData.expires = Date.now() + INVENTORY_ACCESS.sessionDuration;
            localStorage.setItem(INVENTORY_ACCESS.sessionKey, JSON.stringify(sessionData));
            secureLog('Inventory session extended');
        }
    } catch (error) {
        secureLog('Session extension error:', error);
    }
}

// Logout from inventory management
function logoutInventoryAccess() {
    localStorage.removeItem(INVENTORY_ACCESS.sessionKey);
    hideAccessStatusBar();
    showInventoryStatus('🔓 Logged out of inventory management', 'info');
    
    // Switch to orders tab
    const ordersTab = document.getElementById('orders-tab');
    if (ordersTab && window.bootstrap && bootstrap.Tab) {
        const tab = new bootstrap.Tab(ordersTab);
        tab.show();
    }
}

function updateInventoryStats() {
    const validNewTitles = newTitles.filter(title => title.status === 'new').length;
    const duplicates = newTitles.filter(title => 
        booksData.some(book => book.code === title.normalizedISBN)
    ).length;
    
    document.getElementById('currentTitlesCount').textContent = booksData.length;
    document.getElementById('newTitlesCount').textContent = validNewTitles;
    document.getElementById('removeTitlesCount').textContent = titlesToRemove.length;
    document.getElementById('duplicatesCount').textContent = duplicates;
    
    const finalTotal = booksData.length + validNewTitles - titlesToRemove.length;
    const netChange = validNewTitles - titlesToRemove.length;
    
    document.getElementById('totalAfterCount').textContent = Math.max(0, finalTotal);
    document.getElementById('netChangeCount').textContent = netChange >= 0 ? `+${netChange}` : netChange;
    document.getElementById('netChangeCount').className = netChange >= 0 ? 'text-success mb-1' : 'text-danger mb-1';
}

async function processInventoryFile() {
    if (!requireInventoryAccess('file upload')) return;
    
    const file = document.getElementById('inventoryFile').files[0];
    if (!file) return;

    try {
        validateFile(file);
        showInventoryStatus('Processing inventory file...', 'info');
        extendInventorySession(); // Extend session on activity
        
        const arrayBuffer = await file.arrayBuffer();
        let data = [];
        
        if (file.name.toLowerCase().endsWith('.csv')) {
            const text = new TextDecoder().decode(arrayBuffer);
            const parsed = Papa.parse(text, {
                header: true,
                dynamicTyping: true,
                skipEmptyLines: true,
                delimitersToGuess: [',', '\t', '|', ';']
            });
            data = parsed.data;
        } else {
            const workbook = XLSX.read(arrayBuffer);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            data = XLSX.utils.sheet_to_json(worksheet);
        }

        secureLog('Inventory data loaded:', data);
        
        const processedTitles = data.map((row, index) => {
            let isbn = String(row.ISBN || row.isbn || '').trim();
            
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            if (!validateISBN(isbn)) {
                return {
                    originalISBN: isbn,
                    normalizedISBN: '',
                    description: '',
                    status: 'invalid',
                    error: 'Invalid ISBN format'
                };
            }
            
            const normalizedISBN = isbn.replace(/\D/g, '').padStart(13, '0');
            const description = sanitizeText(row.Description || row.description || '');
            
            if (!description) {
                return {
                    originalISBN: isbn,
                    normalizedISBN,
                    description: '',
                    status: 'invalid',
                    error: 'Missing description'
                };
            }
            
            const isDuplicate = booksData.some(book => book.code === normalizedISBN);
            const isNewDuplicate = newTitles.some(title => title.normalizedISBN === normalizedISBN);
            
            return {
                originalISBN: isbn,
                normalizedISBN,
                description,
                status: isDuplicate ? 'duplicate' : (isNewDuplicate ? 'duplicate-new' : 'new'),
                error: null
            };
        });

        newTitles = [...newTitles, ...processedTitles];
        updateNewTitlesTable();
        updateInventoryStats();
        
        const validCount = processedTitles.filter(t => t.status === 'new').length;
        const duplicateCount = processedTitles.filter(t => t.status.includes('duplicate')).length;
        const invalidCount = processedTitles.filter(t => t.status === 'invalid').length;
        
        showInventoryStatus(
            `Processed ${processedTitles.length} titles: ${validCount} new, ${duplicateCount} duplicates, ${invalidCount} invalid`,
            'success'
        );
        
        document.getElementById('downloadInventoryBtn').disabled = newTitles.filter(t => t.status === 'new').length === 0 && titlesToRemove.length === 0;
        document.getElementById('clearNewTitlesBtn').disabled = false;
        
    } catch (error) {
        console.error('Inventory processing error:', error);
        showInventoryStatus('Error processing inventory file: ' + error.message, 'danger');
    }
}

function updateNewTitlesTable() {
    const tbody = document.getElementById('newTitlesBody');
    tbody.innerHTML = '';

    if (newTitles.length === 0) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 4;
        cell.className = 'text-center text-muted';
        cell.textContent = 'No new titles uploaded yet';
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
    }

    newTitles.forEach((title, index) => {
        const tr = document.createElement('tr');
        
        const statusCell = document.createElement('td');
        const badge = document.createElement('span');
        
        switch (title.status) {
            case 'new':
                badge.className = 'badge bg-success';
                badge.textContent = 'New';
                break;
            case 'duplicate':
                badge.className = 'badge bg-warning';
                badge.textContent = 'Duplicate (Existing)';
                break;
            case 'duplicate-new':
                badge.className = 'badge bg-warning';
                badge.textContent = 'Duplicate (In Upload)';
                break;
            case 'invalid':
                badge.className = 'badge bg-danger';
                badge.textContent = 'Invalid';
                break;
        }
        statusCell.appendChild(badge);
        
        if (title.error) {
            statusCell.appendChild(document.createElement('br'));
            const errorSpan = document.createElement('small');
            errorSpan.className = 'text-danger';
            errorSpan.textContent = title.error;
            statusCell.appendChild(errorSpan);
        }
        
        const isbnCell = document.createElement('td');
        isbnCell.textContent = title.normalizedISBN || title.originalISBN;
        
        const descCell = document.createElement('td');
        descCell.textContent = title.description || 'N/A';
        
        const actionCell = document.createElement('td');
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'btn btn-danger btn-sm';
        deleteBtn.innerHTML = '<i class="fas fa-trash"></i>';
        deleteBtn.onclick = () => removeNewTitle(index);
        actionCell.appendChild(deleteBtn);
        
        tr.appendChild(statusCell);
        tr.appendChild(isbnCell);
        tr.appendChild(descCell);
        tr.appendChild(actionCell);
        
        tbody.appendChild(tr);
    });
}

function removeNewTitle(index) {
    newTitles.splice(index, 1);
    updateNewTitlesTable();
    updateInventoryStats();
    
    if (newTitles.length === 0) {
        document.getElementById('downloadInventoryBtn').disabled = true;
        document.getElementById('clearNewTitlesBtn').disabled = true;
        document.getElementById('inventoryFile').value = '';
        document.getElementById('processInventoryBtn').disabled = true;
    }
}

function clearNewTitles() {
    newTitles = [];
    updateNewTitlesTable();
    updateInventoryStats();
    document.getElementById('downloadInventoryBtn').disabled = (newTitles.length === 0 && titlesToRemove.length === 0);
    document.getElementById('clearNewTitlesBtn').disabled = true;
    document.getElementById('inventoryFile').value = '';
    document.getElementById('processInventoryBtn').disabled = true;
    showInventoryStatus('All new titles cleared', 'info');
}

function searchAndMarkForRemoval() {
    if (!requireInventoryAccess('title removal')) return;
    
    const isbnInput = document.getElementById('removeIsbn');
    const rawISBN = sanitizeText(isbnInput.value);
    
    if (!rawISBN) {
        showInventoryStatus('Please enter an ISBN to search for removal', 'warning');
        return;
    }
    
    if (!validateISBN(rawISBN)) {
        showInventoryStatus('Invalid ISBN format. Please enter a valid 10 or 13 digit ISBN.', 'danger');
        return;
    }
    
    const normalizedISBN = rawISBN.replace(/\D/g, '').padStart(13, '0');
    
    // Check if already marked for removal
    if (titlesToRemove.some(title => title.code === normalizedISBN)) {
        showInventoryStatus('This title is already marked for removal', 'warning');
        return;
    }
    
    // Find in current inventory
    const foundBook = booksData.find(book => book.code === normalizedISBN);
    
    if (foundBook) {
        titlesToRemove.push({
            code: normalizedISBN,
            description: foundBook.description
        });
        
        updateRemoveTitlesTable();
        updateInventoryStats();
        isbnInput.value = '';
        extendInventorySession(); // Extend session on activity
        
        document.getElementById('downloadInventoryBtn').disabled = false;
        document.getElementById('clearRemovalsBtn').disabled = false;
        
        showInventoryStatus(`Title marked for removal: ${foundBook.description}`, 'success');
    } else {
        showInventoryStatus(`ISBN ${normalizedISBN} not found in current inventory`, 'danger');
    }
}

async function processBulkRemoval() {
    if (!requireInventoryAccess('bulk removal')) return;
    
    const file = document.getElementById('removeFile').files[0];
    if (!file) return;

    try {
        validateFile(file);
        showInventoryStatus('Processing bulk removal file...', 'info');
        extendInventorySession(); // Extend session on activity
        
        const arrayBuffer = await file.arrayBuffer();
        let data = [];
        
        if (file.name.toLowerCase().endsWith('.txt')) {
            const text = new TextDecoder().decode(arrayBuffer);
            // Split by lines and filter out empty lines
            const lines = text.split(/\r?\n/).filter(line => line.trim());
            data = lines.map(line => ({ ISBN: line.trim() }));
        } else if (file.name.toLowerCase().endsWith('.csv')) {
            const text = new TextDecoder().decode(arrayBuffer);
            const parsed = Papa.parse(text, {
                header: true,
                dynamicTyping: true,
                skipEmptyLines: true,
                delimitersToGuess: [',', '\t', '|', ';']
            });
            data = parsed.data;
        } else {
            const workbook = XLSX.read(arrayBuffer);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            data = XLSX.utils.sheet_to_json(worksheet);
        }

        let addedCount = 0;
        let notFoundCount = 0;
        let duplicateCount = 0;

        data.forEach(row => {
            let isbn = String(row.ISBN || row.isbn || '').trim();
            
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            if (!validateISBN(isbn)) {
                return;
            }
            
            const normalizedISBN = isbn.replace(/\D/g, '').padStart(13, '0');
            
            // Check if already marked for removal
            if (titlesToRemove.some(title => title.code === normalizedISBN)) {
                duplicateCount++;
                return;
            }
            
            // Find in current inventory
            const foundBook = booksData.find(book => book.code === normalizedISBN);
            
            if (foundBook) {
                titlesToRemove.push({
                    code: normalizedISBN,
                    description: foundBook.description
                });
                addedCount++;
            } else {
                notFoundCount++;
            }
        });

        updateRemoveTitlesTable();
        updateInventoryStats();
        
        if (addedCount > 0) {
            document.getElementById('downloadInventoryBtn').disabled = false;
            document.getElementById('clearRemovalsBtn').disabled = false;
        }
        
        showInventoryStatus(
            `Bulk removal processed: ${addedCount} marked for removal, ${notFoundCount} not found, ${duplicateCount} duplicates`,
            'success'
        );
        
    } catch (error) {
        console.error('Bulk removal error:', error);
        showInventoryStatus('Error processing bulk removal file: ' + error.message, 'danger');
    }
}

function updateRemoveTitlesTable() {
    const tbody = document.getElementById('removeTitlesBody');
    tbody.innerHTML = '';

    if (titlesToRemove.length === 0) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 3;
        cell.className = 'text-center text-muted';
        cell.textContent = 'No titles marked for removal';
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
    }

    titlesToRemove.forEach((title, index) => {
        const tr = document.createElement('tr');
        
        const isbnCell = document.createElement('td');
        isbnCell.textContent = title.code;
        
        const descCell = document.createElement('td');
        descCell.textContent = title.description;
        
        const actionCell = document.createElement('td');
        const removeBtn = document.createElement('button');
        removeBtn.className = 'btn btn-outline-success btn-sm';
        removeBtn.innerHTML = '<i class="fas fa-undo"></i> Restore';
        removeBtn.onclick = () => restoreTitle(index);
        actionCell.appendChild(removeBtn);
        
        tr.appendChild(isbnCell);
        tr.appendChild(descCell);
        tr.appendChild(actionCell);
        
        tbody.appendChild(tr);
    });
}

function restoreTitle(index) {
    const restoredTitle = titlesToRemove[index];
    titlesToRemove.splice(index, 1);
    updateRemoveTitlesTable();
    updateInventoryStats();
    
    showInventoryStatus(`Restored: ${restoredTitle.description}`, 'info');
    
    if (titlesToRemove.length === 0 && newTitles.filter(t => t.status === 'new').length === 0) {
        document.getElementById('downloadInventoryBtn').disabled = true;
        document.getElementById('clearRemovalsBtn').disabled = true;
    }
}

function clearRemovals() {
    titlesToRemove = [];
    updateRemoveTitlesTable();
    updateInventoryStats();
    document.getElementById('removeFile').value = '';
    document.getElementById('processBulkRemovalBtn').disabled = true;
    document.getElementById('clearRemovalsBtn').disabled = true;
    
    if (newTitles.filter(t => t.status === 'new').length === 0) {
        document.getElementById('downloadInventoryBtn').disabled = true;
    }
    
    showInventoryStatus('All removals cleared', 'info');
}

function downloadUpdatedInventory() {
    if (!requireInventoryAccess('inventory download')) return;
    
    try {
        const validNewTitles = newTitles.filter(title => title.status === 'new');
        
        if (validNewTitles.length === 0 && titlesToRemove.length === 0) {
            showInventoryStatus('No changes to apply to inventory', 'warning');
            return;
        }
        
        extendInventorySession(); // Extend session on activity
        
        // Start with current inventory
        let mergedInventory = [...booksData];
        
        // Remove titles marked for removal
        if (titlesToRemove.length > 0) {
            const removeISBNs = new Set(titlesToRemove.map(title => title.code));
            mergedInventory = mergedInventory.filter(book => !removeISBNs.has(book.code));
        }
        
        // Add new titles
        validNewTitles.forEach(title => {
            mergedInventory.push({
                code: parseInt(title.normalizedISBN),
                description: title.description
            });
        });
        
        // Sort by ISBN
        mergedInventory.sort((a, b) => a.code - b.code);
        
        const jsonContent = JSON.stringify(mergedInventory, null, 2);
        const blob = new Blob([jsonContent], { type: 'application/json' });
        const url = window.URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        const filename = `data_updated_${now.getFullYear()}_${
            String(now.getMonth() + 1).padStart(2, '0')}_${
            String(now.getDate()).padStart(2, '0')}_${
            String(now.getHours()).padStart(2, '0')}_${
            String(now.getMinutes()).padStart(2, '0')}.json`;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        const changesSummary = [];
        if (validNewTitles.length > 0) changesSummary.push(`${validNewTitles.length} titles added`);
        if (titlesToRemove.length > 0) changesSummary.push(`${titlesToRemove.length} titles removed`);
        
        showInventoryStatus(`📥 Updated inventory downloaded with ${changesSummary.join(', ')}!`, 'success');
        
        secureLog('Inventory update downloaded:', {
            added: validNewTitles.length,
            removed: titlesToRemove.length,
            total: mergedInventory.length
        });
        
    } catch (error) {
        console.error('Download error:', error);
        showInventoryStatus('Error creating updated inventory file: ' + error.message, 'danger');
    }
}

function downloadCurrentInventory() {
    try {
        const jsonContent = JSON.stringify(booksData, null, 2);
        const blob = new Blob([jsonContent], { type: 'application/json' });
        const url = window.URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        const filename = `data_backup_${now.getFullYear()}_${
            String(now.getMonth() + 1).padStart(2, '0')}_${
            String(now.getDate()).padStart(2, '0')}_${
            String(now.getHours()).padStart(2, '0')}_${
            String(now.getMinutes()).padStart(2, '0')}.json`;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        showInventoryStatus('Current inventory backup downloaded!', 'success');
        
    } catch (error) {
        console.error('Backup error:', error);
        showInventoryStatus('Error creating backup file: ' + error.message, 'danger');
    }
}

function showInventoryStatus(message, type) {
    const statusDiv = document.getElementById('inventoryStatus');
    statusDiv.className = `alert alert-${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
}// Global variables
let booksData = [];
let processedOrders = [];
let newTitles = [];
let titlesToRemove = [];
let mergedInventory = [];

// Password protection configuration
const INVENTORY_ACCESS = {
    enabled: true,
    // SHA-256 hash of "admin123" - change this password by generating a new hash
    passwordHash: "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9",
    sessionKey: "clays_inventory_access",
    sessionDuration: 1 * 60 * 60 * 1000 // 1 hour in milliseconds
};

// Security configuration
const SECURITY_CONFIG = {
    maxFileSize: 10 * 1024 * 1024, // 10MB
    allowedMimeTypes: [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ],
    maxDescriptionLength: 200,
    enableDebugLogging: false // PRODUCTION: Set to false for production deployment
};

// Hardcoded Clays POD configuration
const customerConfig = {
    "name": "Clays Ltd",
    "address": {
        "street": "Popson Street",
        "city": "Bungay",
        "region": "Suffolk",
        "country": "UK",
        "postcode": "NR35 1ED",
        "countryCode": "GB"
    },
    "phone": "01986 893 211",
    "type": "Clays POD",
    "csvStructure": ["HDR", "orderNumber", "date", "type", "companyName", "street", "road", "city", "region", "country", "postcode", "countryCode", "phone"],
    "dtlStructure": {
        "marker": "DTL",
        "columns": {
            "isbn": 4,
            "quantity": 5
        }
    }
};

// Security helper functions
function sanitizeText(text) {
    if (typeof text !== 'string') return '';
    
    // Remove potentially dangerous characters and limit length
    return text.replace(/[<>'"&]/g, '')
               .substring(0, SECURITY_CONFIG.maxDescriptionLength)
               .trim();
}

function validateFile(file) {
    if (!file) {
        throw new Error('No file provided');
    }
    
    // Check file size
    if (file.size > SECURITY_CONFIG.maxFileSize) {
        throw new Error(`File size exceeds limit of ${SECURITY_CONFIG.maxFileSize / (1024 * 1024)}MB`);
    }
    
    // Check file type
    if (!SECURITY_CONFIG.allowedMimeTypes.includes(file.type)) {
        throw new Error('Invalid file type. Only Excel files (.xlsx, .xls) are allowed');
    }
    
    // Check file extension as additional validation
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        throw new Error('Invalid file extension. Only .xlsx and .xls files are allowed');
    }
    
    return true;
}

function secureLog(message, data = null) {
    if (SECURITY_CONFIG.enableDebugLogging) {
        console.log(message, data);
    }
}

function validateISBN(isbn) {
    if (typeof isbn !== 'string') return false;
    
    // Remove all non-digits
    const cleanISBN = isbn.replace(/\D/g, '');
    
    // Check if it's a valid length (10 or 13 digits)
    return cleanISBN.length >= 10 && cleanISBN.length <= 13;
}

function validateQuantity(qty) {
    const quantity = parseInt(qty);
    return !isNaN(quantity) && quantity > 0 && quantity <= 10000; // Reasonable upper limit
}

async function fetchData() {
    try {
        secureLog('Attempting to fetch data.json...');
        const response = await fetch('data.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const rawData = await response.json();
        secureLog('Raw data loaded:', rawData);
        
        // Process and normalize the data with security checks
        booksData = rawData.map(item => {
            // Convert numeric codes to strings and pad to 13 digits
            const code = String(item.code).padStart(13, '0');
            
            // Sanitize description
            const description = sanitizeText(item.description || 'No description available');
            
            return {
                code: code,
                description: description,
                setupdate: sanitizeText(item.setupdate || '')
            };
        });
        
        secureLog('Processed books data:', booksData);
        showStatus(`Successfully loaded ${booksData.length} books from inventory`, 'success');
        updateInventoryStats();
        
    } catch (error) {
        console.error('Error fetching data:', error);
        showStatus('Error loading book data. Please ensure data.json exists and is properly formatted.', 'danger');
        booksData = [];
    }
}

document.getElementById('excelFile').addEventListener('change', handleFileSelect);
document.getElementById('inventoryFile').addEventListener('change', function() {
    document.getElementById('processInventoryBtn').disabled = !this.files[0];
});
document.getElementById('removeFile').addEventListener('change', function() {
    document.getElementById('processBulkRemovalBtn').disabled = !this.files[0];
});
document.getElementById('removeIsbn').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        searchAndMarkForRemoval();
    }
});
document.getElementById('isbnSearch').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        searchISBN();
    }
});

function searchISBN() {
    const isbnInput = document.getElementById('isbnSearch');
    const resultDiv = document.getElementById('isbnResult');
    const alertDiv = document.getElementById('isbnResultAlert');
    const contentDiv = document.getElementById('isbnResultContent');
    
    const rawISBN = sanitizeText(isbnInput.value);
    
    if (!rawISBN) {
        showISBNResult('Please enter an ISBN to search', 'warning');
        return;
    }
    
    // Validate and normalize ISBN
    if (!validateISBN(rawISBN)) {
        showISBNResult('Invalid ISBN format. Please enter a valid 10 or 13 digit ISBN.', 'danger');
        return;
    }
    
    // Normalize ISBN (remove non-digits and pad)
    let normalizedISBN = rawISBN.replace(/\D/g, '').padStart(13, '0');
    
    // Search in the books database
    const foundBook = booksData.find(book => book.code === normalizedISBN);
    
    if (foundBook) {
        showISBNResult(`
            <strong>✓ Available</strong><br>
            <strong>ISBN:</strong> ${normalizedISBN}<br>
            <strong>Title:</strong> ${foundBook.description}<br>
            ${foundBook.setupdate ? `<strong>Setup Date:</strong> ${foundBook.setupdate}` : ''}
        `, 'success');
    } else {
        showISBNResult(`
            <strong>✗ Not Available</strong><br>
            <strong>ISBN:</strong> ${normalizedISBN}<br>
            This title is not available in the current inventory.
        `, 'danger');
    }
}

function showISBNResult(content, type) {
    const resultDiv = document.getElementById('isbnResult');
    const alertDiv = document.getElementById('isbnResultAlert');
    const contentDiv = document.getElementById('isbnResultContent');
    
    alertDiv.className = `alert alert-${type}`;
    contentDiv.innerHTML = content;
    resultDiv.style.display = 'block';
    
    // Auto-hide success/info results after 10 seconds
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            resultDiv.style.display = 'none';
        }, 10000);
    }
}

async function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    const orderRef = sanitizeText(document.getElementById('orderRef').value);
    const orderRefWarning = document.getElementById('orderRefWarning');
    
    if (!orderRef) {
        orderRefWarning.style.display = 'block';
        showStatus('Please enter an order reference before uploading a file', 'warning');
        document.getElementById('excelFile').value = '';
        return;
    }
    
    orderRefWarning.style.display = 'none';

    try {
        // Validate file security
        validateFile(file);
        
        showStatus('Processing file...', 'info');
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        secureLog('Excel data loaded:', excelData);

        // Create a map for faster lookup
        const booksMap = new Map();
        booksData.forEach(item => {
            booksMap.set(item.code, item);
        });

        secureLog('Books map created:', booksMap);

        processedOrders = excelData.map((row, index) => {
            secureLog('Processing Excel row:', row);
            
            // Handle ISBN with security validation
            let isbn = String(row.ISBN || row.isbn || '').trim();
            
            // Handle scientific notation in Excel
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            // Validate and normalize ISBN
            if (!validateISBN(isbn)) {
                secureLog('Invalid ISBN detected:', isbn);
                isbn = '';
            } else {
                isbn = isbn.replace(/\D/g, '').padStart(13, '0');
            }
            
            // Handle quantity with validation
            const rawQuantity = row.Qty || row.qty || row.Quantity || row.quantity || 0;
            const quantity = validateQuantity(rawQuantity) ? parseInt(rawQuantity) : 0;
            
            secureLog('Normalized ISBN:', isbn, 'Quantity:', quantity);
            
            // Look up book in inventory
            const stockItem = booksMap.get(isbn);
            secureLog('Stock item found:', stockItem);
            
            return {
                lineNumber: String(index + 1).padStart(3, '0'),
                orderRef: sanitizeText(orderRef),
                isbn,
                description: stockItem?.description || 'Not Found',
                quantity,
                available: !!stockItem,
                setupDate: stockItem?.setupdate || ''
            };
        });

        secureLog('Processed orders:', processedOrders);
        updatePreviewTable();
        
        const availableCount = processedOrders.filter(order => order.available).length;
        const totalCount = processedOrders.length;
        showStatus(`Data loaded successfully! ${availableCount}/${totalCount} items found in inventory.`, 'success');
        enableButtons(true);
        
    } catch (error) {
        console.error('Processing error:', error);
        showStatus('Error processing file: ' + error.message, 'danger');
        enableButtons(false);
        // Clear file input on error
        document.getElementById('excelFile').value = '';
    }
}

function updatePreviewTable() {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';

    if (processedOrders.length === 0) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 7;
        cell.className = 'text-center';
        cell.textContent = 'No data loaded';
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
    }

    const showOnlyUnavailable = document.getElementById('showOnlyUnavailable').checked;
    let filteredOrders = processedOrders;

    if (showOnlyUnavailable) {
        filteredOrders = processedOrders.filter(order => !order.available);
    }

    if (filteredOrders.length === 0 && showOnlyUnavailable) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 7;
        cell.className = 'text-center text-success';
        cell.textContent = 'All items are available in inventory!';
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
    }

    filteredOrders.forEach((order, index) => {
        const originalIndex = processedOrders.indexOf(order);
        
        const tr = document.createElement('tr');
        
        // Create cells using secure methods
        const checkboxCell = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'row-checkbox';
        checkbox.dataset.index = originalIndex;
        checkboxCell.appendChild(checkbox);
        
        const actionCell = document.createElement('td');
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'btn btn-danger btn-sm';
        deleteBtn.onclick = () => deleteRow(originalIndex);
        deleteBtn.innerHTML = '<i class="fas fa-trash"></i>';
        actionCell.appendChild(deleteBtn);
        
        const lineCell = document.createElement('td');
        lineCell.textContent = order.lineNumber;
        
        const isbnCell = document.createElement('td');
        isbnCell.textContent = order.isbn;
        
        const descCell = document.createElement('td');
        descCell.textContent = order.description; // Safe from XSS
        
        const qtyCell = document.createElement('td');
        qtyCell.textContent = order.quantity;
        
        const statusCell = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = `badge ${order.available ? 'bg-success' : 'bg-danger'}`;
        badge.textContent = order.available ? 'Available' : 'Not Found';
        statusCell.appendChild(badge);
        
        tr.appendChild(checkboxCell);
        tr.appendChild(actionCell);
        tr.appendChild(lineCell);
        tr.appendChild(isbnCell);
        tr.appendChild(descCell);
        tr.appendChild(qtyCell);
        tr.appendChild(statusCell);
        
        tbody.appendChild(tr);
    });
}

function formatDate() {
    const now = new Date();
    return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
}

function deleteRow(index) {
    processedOrders.splice(index, 1);
    processedOrders = processedOrders.map((order, idx) => ({
        ...order,
        lineNumber: String(idx + 1).padStart(3, '0')
    }));
    updatePreviewTable();
    showStatus(`Row ${index + 1} deleted`, 'info');
    enableButtons(processedOrders.length > 0);
}

function clearAll() {
    processedOrders = [];
    updatePreviewTable();
    document.getElementById('excelFile').value = '';
    document.getElementById('orderRef').value = '';
    showStatus('All data cleared', 'info');
    enableButtons(false);
}

function downloadTemplate() {
    const link = document.createElement('a');
    link.href = 'order_template.xlsx';
    link.download = 'order_template.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function downloadCsv() {
    if (processedOrders.length === 0) {
        showStatus('No data to download', 'warning');
        return;
    }

    try {
        const formattedDate = formatDate();
        const orderRef = sanitizeText(document.getElementById('orderRef').value);
    
        const csvRow = customerConfig.csvStructure.map(field => {
            switch (field) {
                case 'HDR':
                    return 'HDR';
                case 'orderNumber':
                    return orderRef;
                case 'date':
                    return formattedDate;
                case 'code':
                    return customerConfig.headerCode || '';
                case 'type':
                    return customerConfig.type || '';
                case 'companyName':
                    return customerConfig.name;
                case 'phone':
                    return customerConfig.phone;
                default:
                    return customerConfig.address[field] || '';
            }
        });
    
        const finalCsvRow = [...csvRow, ''];
        const csvContent = [finalCsvRow];
    
        processedOrders.forEach(order => {
            const dtlRow = Array(customerConfig.csvStructure.length).fill('');
            dtlRow[0] = 'DTL';
            dtlRow[1] = orderRef;
            dtlRow[2] = order.lineNumber;
            dtlRow[3] = order.isbn;
            dtlRow[4] = order.quantity.toString();
            dtlRow.push('');
            csvContent.push(dtlRow);
        });
    
        const csv = Papa.unparse(csvContent);
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        const filename = `pod_order_${now.getFullYear()}_${
            String(now.getMonth() + 1).padStart(2, '0')}_${
            String(now.getDate()).padStart(2, '0')}_${
            String(now.getHours()).padStart(2, '0')}_${
            String(now.getMinutes()).padStart(2, '0')}_${
            String(now.getSeconds()).padStart(2, '0')}.csv`;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    
        showStatus('CSV downloaded successfully!', 'success');
    } catch (error) {
        console.error('Download error:', error);
        showStatus('Error creating CSV file: ' + error.message, 'danger');
    }
}

function enableButtons(enabled) {
    document.getElementById('clearBtn').disabled = !enabled;
    document.getElementById('downloadBtn').disabled = !enabled;
    document.getElementById('deleteSelectedBtn').disabled = !enabled;
    document.getElementById('copyBtn').disabled = !enabled;
    document.getElementById('selectAll').checked = false;
}

function copyTableToClipboard() {
    if (processedOrders.length === 0) return;

    let tableHtml = `
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr style="background-color: #f8f9fa;">
                    <th style="padding: 8px; border: 1px solid #dee2e6;">Line No</th>
                    <th style="padding: 8px; border: 1px solid #dee2e6;">ISBN</th>
                    <th style="padding: 8px; border: 1px solid #dee2e6;">Description</th>
                    <th style="padding: 8px; border: 1px solid #dee2e6;">Quantity</th>
                    <th style="padding: 8px; border: 1px solid #dee2e6;">Status</th>
                </tr>
            </thead>
            <tbody>
    `;

    processedOrders.forEach(order => {
        // Escape HTML in data for safe insertion
        const safeDescription = order.description.replace(/[<>&"']/g, function(match) {
            const escapeMap = {
                '<': '&lt;',
                '>': '&gt;',
                '&': '&amp;',
                '"': '&quot;',
                "'": '&#39;'
            };
            return escapeMap[match];
        });
        
        tableHtml += `
            <tr>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${order.lineNumber}</td>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${order.isbn}</td>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${safeDescription}</td>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${order.quantity}</td>
                <td style="padding: 8px; border: 1px solid #dee2e6; ${order.available ? 'color: green;' : 'color: red;'}">${order.available ? 'Available' : 'Not Found'}</td>
            </tr>
        `;
    });

    tableHtml += '</tbody></table>';

    const blob = new Blob([tableHtml], { type: 'text/html' });
    const clipboardItem = new ClipboardItem({ 'text/html': blob });
    
    navigator.clipboard.write([clipboardItem]).then(() => {
        showStatus('Table copied to clipboard', 'success');
    }).catch(err => {
        console.error('Failed to copy:', err);
        showStatus('Failed to copy table', 'danger');
    });
}

function toggleAllCheckboxes() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const selectAllCheckbox = document.getElementById('selectAll');
    checkboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
    });
}

function deleteSelected() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    const indices = Array.from(checkboxes)
        .map(checkbox => parseInt(checkbox.dataset.index))
        .sort((a, b) => b - a);

    indices.forEach(index => {
        processedOrders.splice(index, 1);
    });

    processedOrders = processedOrders.map((order, idx) => ({
        ...order,
        lineNumber: String(idx + 1).padStart(3, '0')
    }));

    updatePreviewTable();
    showStatus(`${indices.length} rows deleted`, 'info');
    enableButtons(processedOrders.length > 0);
}

function toggleTableFilter() {
    updatePreviewTable();
    document.getElementById('selectAll').checked = false;
    
    const showOnlyUnavailable = document.getElementById('showOnlyUnavailable').checked;
    const unavailableCount = processedOrders.filter(order => !order.available).length;
    
    if (showOnlyUnavailable) {
        showStatus(`Showing ${unavailableCount} unavailable items`, 'info');
    } else {
        showStatus(`Showing all ${processedOrders.length} items`, 'info');
    }
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.className = `alert alert-${type}`;
    statusDiv.textContent = message; // Safe from XSS
    statusDiv.style.display = 'block';
    
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
}

// Initialize the application
secureLog('Initializing application...');
fetchData();