// Global variables
let booksData = [];
let processedOrders = [];

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
        
    } catch (error) {
        console.error('Error fetching data:', error);
        showStatus('Error loading book data. Please ensure data.json exists and is properly formatted.', 'danger');
        booksData = [];
    }
}

document.getElementById('excelFile').addEventListener('change', handleFileSelect);
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