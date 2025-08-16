function downloadJsonTemplate() {
    const templateData = [
        {
            "ISBN": 9780140175936,
            "Master Order ID": "SA1657",
            "TITLE": "CLEOPATRA'S SISTER   (11)",
            "Trim Height": 198,
            "Trim Width": 129,
            "Bind Style": "PU/2",
            "Extent": 288,
            "Paper Desc": "Holmen Bulky 52 gsm",
            "Cover Spec Code 1": "C400P2",
            "Cover Spine": 18,
            "Packing": "Pack (64) (8) Base.",
            "Status": "POD Ready"
        },
        {
            "ISBN": 9780140256932,
            "Master Order ID": "SA1659",
            "TITLE": "BEYOND THE BLUE MOUNTAINS (07)",
            "Trim Height": 198,
            "Trim Width": 129,
            "Bind Style": "PU/2",
            "Extent": 160,
            "Paper Desc": "Holmen Bulky 52 gsm",
            "Cover Spec Code 1": "C400P2",
            "Cover Spine": 11,
            "Packing": "Pack (104) (8) Base.",
            "Status": "MPI"
        }
    ];
    
    const jsonContent = JSON.stringify(templateData, null, 2);
    const blob = new Blob([jsonContent], { type: 'application/json' });
    const url = window.URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'new_titles_template.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    showInventoryStatus('JSON template downloaded!', 'success');
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
    sessionKey: "clays_repo_access",
    sessionDuration: 1 * 60 * 60 * 1000 // 1 hour in milliseconds
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

// Security configuration
const SECURITY_CONFIG = {
    maxFileSize: 10 * 1024 * 1024, // 10MB
    allowedMimeTypes: [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
        'application/csv',
        'application/json',
        'text/json',
        'text/plain' // For .txt files in removal
    ],
    maxDescriptionLength: 200,
    enableDebugLogging: false // PRODUCTION: Set to false for production deployment
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
        throw new Error('Invalid file type. Only Excel (.xlsx, .xls), CSV (.csv), and JSON (.json) files are allowed');
    }
    
    // Check file extension as additional validation
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv') && !fileName.endsWith('.json')) {
        throw new Error('Invalid file extension. Only .xlsx, .xls, .csv, and .json files are allowed');
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
            // Convert numeric ISBNs to strings and pad to 13 digits
            const isbn = String(item.ISBN).padStart(13, '0');
            
            // Sanitize title and master order ID
            const title = sanitizeText(item.TITLE || 'No title available');
            const masterOrderId = sanitizeText(item['Master Order ID'] || '');
            const status = sanitizeText(item.Status || 'POD Ready'); // Default to POD Ready if missing
            
            return {
                isbn: isbn,
                title: title,
                masterOrderId: masterOrderId,
                status: status,
                // Keep original structure for future development
                trimHeight: item['Trim Height'],
                trimWidth: item['Trim Width'],
                bindStyle: sanitizeText(item['Bind Style'] || ''),
                extent: item.Extent,
                paperDesc: sanitizeText(item['Paper Desc'] || ''),
                coverSpecCode1: sanitizeText(item['Cover Spec Code 1'] || ''),
                coverSpine: item['Cover Spine'],
                packing: sanitizeText(item.Packing || ''),
                // Legacy compatibility
                code: isbn,
                description: title,
                setupdate: sanitizeText(item.setupdate || '')
            };
        });
        
        secureLog('Processed books data:', booksData);
        showStatus(`Successfully loaded ${booksData.length} books from repo`, 'success');
        updateInventoryStats();
        
    } catch (error) {
        console.error('Error fetching data:', error);
        showStatus('Error loading book data. Please ensure data.json exists and is properly formatted.', 'danger');
        booksData = [];
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
        let excelData = [];
        
        // Handle different file types
        if (file.name.toLowerCase().endsWith('.csv')) {
            const text = new TextDecoder().decode(arrayBuffer);
            const parsed = Papa.parse(text, {
                header: true,
                dynamicTyping: true,
                skipEmptyLines: true,
                delimitersToGuess: [',', '\t', '|', ';']
            });
            excelData = parsed.data;
        } else {
            // Handle Excel files
            const workbook = XLSX.read(arrayBuffer);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(worksheet);
        }

        secureLog('File data loaded:', excelData);

        // Create a map for faster lookup (supporting both ISBN and Master Order ID)
        const booksMap = new Map();
        const masterOrderMap = new Map();
        booksData.forEach(item => {
            booksMap.set(item.isbn, item);
            if (item.masterOrderId) {
                masterOrderMap.set(item.masterOrderId.toLowerCase(), item);
            }
        });

        secureLog('Books map created:', booksMap);

        processedOrders = excelData.map((row, index) => {
            secureLog('Processing row:', row);
            
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
            
            // Look up book in repo
            const stockItem = booksMap.get(isbn);
            secureLog('Stock item found:', stockItem);
            
            return {
                lineNumber: String(index + 1).padStart(3, '0'),
                orderRef: sanitizeText(orderRef),
                isbn,
                description: stockItem?.title || 'Not Found',
                quantity,
                available: !!stockItem,
                status: stockItem?.status || 'Not Available',
                masterOrderId: stockItem?.masterOrderId || ''
            };
        });

        secureLog('Processed orders:', processedOrders);
        updatePreviewTable();
        
        const validCount = processedOrders.filter(order => order.available).length;
        const totalCount = processedOrders.length;
        const podReadyCount = processedOrders.filter(order => order.available && order.status === 'POD Ready').length;
        const mpiCount = processedOrders.filter(order => order.available && order.status === 'MPI').length;
        const notAvailableCount = processedOrders.filter(order => !order.available).length;
        
        showStatus(`Data loaded successfully! ${validCount}/${totalCount} items found in repo (${podReadyCount} POD Ready, ${mpiCount} MPI, ${notAvailableCount} Not Available).`, 'success');
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

    const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
    const showNotAvailable = document.getElementById('showNotAvailable') ? document.getElementById('showNotAvailable').checked : false;
    let filteredOrders = processedOrders;

    if (printAsMiscellaneous) {
        // Show only MPI items
        filteredOrders = processedOrders.filter(order => order.available && order.status === 'MPI');
    } else if (showNotAvailable) {
        // Show only items not in repo
        filteredOrders = processedOrders.filter(order => !order.available);
    }

    if (filteredOrders.length === 0) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 7;
        cell.className = 'text-center text-success';
        
        if (printAsMiscellaneous) {
            cell.textContent = 'No MPI items found!';
        } else if (showNotAvailable) {
            cell.textContent = 'All items are available in repo!';
        } else {
            cell.textContent = 'No data to display';
        }
        
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
        
        // Updated status logic based on new Status field
        if (!order.available) {
            badge.className = 'badge bg-danger';
            badge.textContent = 'Not Available';
        } else {
            switch (order.status) {
                case 'POD Ready':
                    badge.className = 'badge bg-success';
                    badge.textContent = 'POD Ready';
                    break;
                case 'MPI':
                    badge.className = 'badge bg-warning text-dark';
                    badge.textContent = 'MPI';
                    break;
                default:
                    badge.className = 'badge bg-info';
                    badge.textContent = order.status;
                    break;
            }
        }
        
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

function searchISBN() {
    const isbnInput = document.getElementById('isbnSearch');
    const resultDiv = document.getElementById('isbnResult');
    const alertDiv = document.getElementById('isbnResultAlert');
    const contentDiv = document.getElementById('isbnResultContent');
    
    const rawInput = sanitizeText(isbnInput.value);
    
    if (!rawInput) {
        showISBNResult('Please enter an ISBN or Master Order ID to search', 'warning');
        return;
    }
    
    let foundBook = null;
    let searchType = '';
    
    // First try to search as ISBN
    if (validateISBN(rawInput)) {
        const normalizedISBN = rawInput.replace(/\D/g, '').padStart(13, '0');
        foundBook = booksData.find(book => book.isbn === normalizedISBN);
        searchType = 'ISBN';
    }
    
    // If not found and doesn't look like ISBN, try Master Order ID
    if (!foundBook) {
        foundBook = booksData.find(book => 
            book.masterOrderId && book.masterOrderId.toLowerCase() === rawInput.toLowerCase()
        );
        searchType = 'Master Order ID';
    }
    
    if (foundBook) {
        showISBNResult(`
            <strong>✓ Available</strong><br>
            <strong>ISBN:</strong> ${foundBook.isbn}<br>
            <strong>Title:</strong> ${foundBook.title}<br>
            <strong>Master Order ID:</strong> ${foundBook.masterOrderId || 'N/A'}<br>
            <strong>Status:</strong> <span class="badge ${getStatusBadgeClass(foundBook.status)}">${foundBook.status}</span><br>
            <small class="text-muted">Found by ${searchType}</small>
        `, 'success');
    } else {
        // Determine what we were searching for in the error message
        const searchTypeMsg = validateISBN(rawInput) ? 'ISBN' : 'Master Order ID';
        showISBNResult(`
            <strong>✗ Not Available</strong><br>
            <strong>${searchTypeMsg}:</strong> ${rawInput}<br>
            This title is not available in the current repo.
        `, 'danger');
    }
}

function getStatusBadgeClass(status) {
    switch (status) {
        case 'POD Ready':
            return 'bg-success';
        case 'MPI':
            return 'bg-warning text-dark';
        default:
            return 'bg-info';
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
    // Create a link element
    const link = document.createElement('a');
    
    // Set the href to the template file path
    link.href = 'order_template.xlsx';
    
    // Set the download attribute to force download with a specific filename
    link.download = 'order_template.xlsx';
    
    // Temporarily add the link to the document
    document.body.appendChild(link);
    
    // Trigger the download
    link.click();
    
    // Remove the link from the document
    document.body.removeChild(link);
}

function downloadNewTitlesTemplate() {
    // Create a link element for Excel template
    const link = document.createElement('a');
    
    // Set the href to the new titles template file path
    link.href = 'new_titles.xlsx';
    
    // Set the download attribute to force download with a specific filename
    link.download = 'new_titles.xlsx';
    
    // Temporarily add the link to the document
    document.body.appendChild(link);
    
    // Trigger the download
    link.click();
    
    // Remove the link from the document
    document.body.removeChild(link);
}

function downloadJsonTemplate() {
    const templateData = [
        {
            "ISBN": 9780140064810,
            "Master Order ID": "SA1693",
            "TITLE": "NEXT TO NATURE, ART (10)",
            "Trim Height": 198,
            "Trim Width": 129,
            "Bind Style": "PU/2",
            "Extent": 192,
            "Paper Desc": "Holmen Bulky 52 gsm",
            "Cover Spec Code 1": "C400P2",
            "Cover Spine": 13,
            "Packing": "Pack (88) (8) Base."
        },
        {
            "ISBN": 9780140119329,
            "Master Order ID": "SA2055",
            "TITLE": "PASSING ON (14)",
            "Trim Height": 198,
            "Trim Width": 129,
            "Bind Style": "PU/2",
            "Extent": 224,
            "Paper Desc": "Holmen Bulky 52 gsm",
            "Cover Spec Code 1": "C400P2",
            "Cover Spine": 14,
            "Packing": "Pack (80) (8) Base."
        }
    ];
    
    const jsonContent = JSON.stringify(templateData, null, 2);
    const blob = new Blob([jsonContent], { type: 'application/json' });
    const url = window.URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'new_titles_template.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    showInventoryStatus('JSON template downloaded!', 'success');
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
    
    const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
    const showNotAvailable = document.getElementById('showNotAvailable') ? document.getElementById('showNotAvailable').checked : false;
    
    // Ensure only one filter is active at a time
    if (printAsMiscellaneous && showNotAvailable) {
        // If both are checked, uncheck the other one
        if (event.target.id === 'showOnlyUnavailable') {
            document.getElementById('showNotAvailable').checked = false;
        } else {
            document.getElementById('showOnlyUnavailable').checked = false;
        }
        // Re-run the function with updated state
        setTimeout(toggleTableFilter, 10);
        return;
    }
    
    const mpiCount = processedOrders.filter(order => order.available && order.status === 'MPI').length;
    const notAvailableCount = processedOrders.filter(order => !order.available).length;
    
    if (printAsMiscellaneous) {
        showStatus(`Showing ${mpiCount} MPI items`, 'info');
    } else if (showNotAvailable) {
        showStatus(`Showing ${notAvailableCount} not available items`, 'info');
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

// Password protection functions
function showInventoryTab() {
    // Check if repo access is required and granted
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
            showInventoryStatus('✅ Repo management access granted', 'success');
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

function updateInventoryStats() {
    const validNewTitles = newTitles.filter(title => title.status === 'new').length;
    const duplicates = newTitles.filter(title => 
        booksData.some(book => book.isbn === title.normalizedISBN)
    ).length;
    
    // Count current repo titles by status
    const podReadyTitles = booksData.filter(book => book.status === 'POD Ready').length;
    const mpiTitles = booksData.filter(book => book.status === 'MPI').length;
    
    document.getElementById('currentTitlesCount').textContent = booksData.length;
    document.getElementById('podReadyCount').textContent = podReadyTitles;
    document.getElementById('mpiCount').textContent = mpiTitles;
    document.getElementById('newTitlesCount').textContent = validNewTitles;
    document.getElementById('removeTitlesCount').textContent = titlesToRemove.length;
    document.getElementById('duplicatesCount').textContent = duplicates;
}

// Check if user has current repo access
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

// Request repo access with password
function requestInventoryAccess() {
    // Create a proper password input dialog
    const passwordDialog = document.createElement('div');
    passwordDialog.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.5);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 9999;
    `;
    
    passwordDialog.innerHTML = `
        <div style="background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); min-width: 300px;">
            <h5 style="margin-bottom: 20px; text-align: center;">🔐 Repo Management Access</h5>
            <div style="margin-bottom: 15px;">
                <label style="display: block; margin-bottom: 5px; font-weight: bold;">Enter Password:</label>
                <input type="password" id="passwordInput" style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;" placeholder="Password">
            </div>
            <div style="display: flex; gap: 10px; justify-content: flex-end;">
                <button id="cancelBtn" style="padding: 8px 16px; border: 1px solid #ccc; background: #f8f9fa; border-radius: 4px; cursor: pointer;">Cancel</button>
                <button id="submitBtn" style="padding: 8px 16px; border: none; background: #007bff; color: white; border-radius: 4px; cursor: pointer;">Submit</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(passwordDialog);
    
    const passwordInput = document.getElementById('passwordInput');
    const submitBtn = document.getElementById('submitBtn');
    const cancelBtn = document.getElementById('cancelBtn');
    
    // Focus on password input
    passwordInput.focus();
    
    return new Promise((resolve) => {
        const cleanup = () => {
            document.body.removeChild(passwordDialog);
        };
        
        const handleSubmit = async () => {
            const password = passwordInput.value;
            cleanup();
            
            if (!password) {
                showInventoryStatus('Access cancelled', 'info');
                resolve(false);
                return;
            }
            
            const success = await validatePasswordAndGrantAccess(password);
            resolve(success);
        };
        
        const handleCancel = () => {
            cleanup();
            showInventoryStatus('Access cancelled', 'info');
            resolve(false);
        };
        
        // Event listeners
        submitBtn.addEventListener('click', handleSubmit);
        cancelBtn.addEventListener('click', handleCancel);
        
        // Allow Enter key to submit
        passwordInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                handleSubmit();
            }
        });
        
        // Allow Escape key to cancel
        passwordDialog.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                handleCancel();
            }
        });
    });
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
            
            secureLog('Repo access granted');
            return true;
        } else {
            alert("❌ Incorrect password. Access denied.");
            secureLog('Repo access denied - incorrect password');
            return false;
        }
    } catch (error) {
        console.error('Password validation error:', error);
        alert("❌ Error validating password. Please try again.");
        return false;
    }
}

// Check if repo functions should be protected
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
            showInventoryStatus(`⏰ Repo access expires in ${minutesLeft} minutes`, 'warning');
        } 
        // Session expired
        else if (timeLeft <= 0) {
            localStorage.removeItem(INVENTORY_ACCESS.sessionKey);
            hideAccessStatusBar();
            showInventoryStatus('🔒 Repo access expired. Please re-authenticate.', 'danger');
            
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
            secureLog('Repo session extended');
        }
    } catch (error) {
        secureLog('Session extension error:', error);
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

// Logout from repo management
function logoutInventoryAccess() {
    localStorage.removeItem(INVENTORY_ACCESS.sessionKey);
    hideAccessStatusBar();
    showInventoryStatus('🔓 Logged out of repo management', 'info');
    
    // Switch to orders tab
    const ordersTab = document.getElementById('orders-tab');
    if (ordersTab && window.bootstrap && bootstrap.Tab) {
        const tab = new bootstrap.Tab(ordersTab);
        tab.show();
    }
}

async function processInventoryFile() {
    if (!requireInventoryAccess('file upload')) return;
    
    const file = document.getElementById('inventoryFile').files[0];
    if (!file) return;

    try {
        // Only validate JSON files for repo management
        if (!file.name.toLowerCase().endsWith('.json')) {
            throw new Error('Only JSON files are supported for repo management');
        }
        
        if (file.size > SECURITY_CONFIG.maxFileSize) {
            throw new Error(`File size exceeds limit of ${SECURITY_CONFIG.maxFileSize / (1024 * 1024)}MB`);
        }
        
        showInventoryStatus('Processing JSON repo file...', 'info');
        extendInventorySession(); // Extend session on activity
        
        const arrayBuffer = await file.arrayBuffer();
        const text = new TextDecoder().decode(arrayBuffer);
        let data = [];
        
        try {
            const jsonData = JSON.parse(text);
            if (Array.isArray(jsonData)) {
                data = jsonData;
            } else {
                throw new Error('JSON file must contain an array of book objects');
            }
        } catch (jsonError) {
            throw new Error('Invalid JSON format: ' + jsonError.message);
        }

        secureLog('JSON repo data loaded:', data);
        
        const processedTitles = data.map((row, index) => {
            // Validate JSON structure
            if (!row.hasOwnProperty('ISBN') || !row.hasOwnProperty('TITLE')) {
                return {
                    originalISBN: row.ISBN || 'missing',
                    normalizedISBN: '',
                    title: '',
                    status: 'invalid',
                    error: 'Missing required fields (ISBN, TITLE)',
                    isFullFormat: false,
                    fullData: null
                };
            }
            
            let isbn = String(row.ISBN || '').trim();
            
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            if (!validateISBN(isbn)) {
                return {
                    originalISBN: isbn,
                    normalizedISBN: '',
                    title: '',
                    status: 'invalid',
                    error: 'Invalid ISBN format',
                    isFullFormat: false,
                    fullData: null
                };
            }
            
            const normalizedISBN = isbn.replace(/\D/g, '').padStart(13, '0');
            const title = sanitizeText(row.TITLE || '');
            
            if (!title) {
                return {
                    originalISBN: isbn,
                    normalizedISBN,
                    title: '',
                    status: 'invalid',
                    error: 'Missing title',
                    isFullFormat: false,
                    fullData: null
                };
            }
            
            const isDuplicate = booksData.some(book => book.isbn === normalizedISBN);
            const isNewDuplicate = newTitles.some(title => title.normalizedISBN === normalizedISBN);
            
            return {
                originalISBN: isbn,
                normalizedISBN,
                title,
                status: isDuplicate ? 'duplicate' : (isNewDuplicate ? 'duplicate-new' : 'new'),
                error: null,
                isFullFormat: true,
                fullData: row
            };
        });

        newTitles = [...newTitles, ...processedTitles];
        updateNewTitlesTable();
        updateInventoryStats();
        
        const validCount = processedTitles.filter(t => t.status === 'new').length;
        const duplicateCount = processedTitles.filter(t => t.status.includes('duplicate')).length;
        const invalidCount = processedTitles.filter(t => t.status === 'invalid').length;
        
        showInventoryStatus(
            `Processed ${processedTitles.length} JSON records: ${validCount} new, ${duplicateCount} duplicates, ${invalidCount} invalid`,
            'success'
        );
        
        document.getElementById('downloadInventoryBtn').disabled = newTitles.filter(t => t.status === 'new').length === 0 && titlesToRemove.length === 0;
        document.getElementById('clearNewTitlesBtn').disabled = false;
        
    } catch (error) {
        console.error('JSON processing error:', error);
        showInventoryStatus('Error processing JSON file: ' + error.message, 'danger');
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
                badge.textContent = 'New (JSON)';
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
        descCell.textContent = title.title || 'N/A';
        
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
        document.getElementById('downloadInventoryBtn').disabled = titlesToRemove.length === 0;
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
    if (titlesToRemove.some(title => title.isbn === normalizedISBN)) {
        showInventoryStatus('This title is already marked for removal', 'warning');
        return;
    }
    
    // Find in current repo
    const foundBook = booksData.find(book => book.isbn === normalizedISBN);
    
    if (foundBook) {
        titlesToRemove.push({
            isbn: normalizedISBN,
            title: foundBook.title
        });
        
        updateRemoveTitlesTable();
        updateInventoryStats();
        isbnInput.value = '';
        extendInventorySession(); // Extend session on activity
        
        document.getElementById('downloadInventoryBtn').disabled = false;
        document.getElementById('clearRemovalsBtn').disabled = false;
        
        showInventoryStatus(`Title marked for removal: ${foundBook.title}`, 'success');
    } else {
        showInventoryStatus(`ISBN ${normalizedISBN} not found in current repo`, 'danger');
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
        } else if (file.name.toLowerCase().endsWith('.json')) {
            const text = new TextDecoder().decode(arrayBuffer);
            try {
                const jsonData = JSON.parse(text);
                if (Array.isArray(jsonData)) {
                    data = jsonData;
                } else {
                    throw new Error('JSON file must contain an array of objects with ISBN field');
                }
            } catch (jsonError) {
                throw new Error('Invalid JSON format: ' + jsonError.message);
            }
        } else {
            throw new Error('Only JSON and TXT files are supported for bulk removal');
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
            if (titlesToRemove.some(title => title.isbn === normalizedISBN)) {
                duplicateCount++;
                return;
            }
            
            // Find in current repo
            const foundBook = booksData.find(book => book.isbn === normalizedISBN);
            
            if (foundBook) {
                titlesToRemove.push({
                    isbn: normalizedISBN,
                    title: foundBook.title
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
        isbnCell.textContent = title.isbn;
        
        const descCell = document.createElement('td');
        descCell.textContent = title.title;
        
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
    
    showInventoryStatus(`Restored: ${restoredTitle.title}`, 'info');
    
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
    if (!requireInventoryAccess('repo download')) return;
    
    try {
        const validNewTitles = newTitles.filter(title => title.status === 'new');
        
        if (validNewTitles.length === 0 && titlesToRemove.length === 0) {
            showInventoryStatus('No changes to apply to repo', 'warning');
            return;
        }
        
        extendInventorySession(); // Extend session on activity
        
        // Start with current repo
        let mergedInventory = [...booksData];
        
        // Remove titles marked for removal
        if (titlesToRemove.length > 0) {
            const removeISBNs = new Set(titlesToRemove.map(title => title.isbn));
            mergedInventory = mergedInventory.filter(book => !removeISBNs.has(book.isbn));
        }
        
        // Add new titles - all are full JSON format now
        validNewTitles.forEach(title => {
            if (title.fullData) {
                // Use the complete JSON structure from the uploaded file
                const fullRecord = { ...title.fullData };
                // Ensure ISBN is normalized
                fullRecord.ISBN = parseInt(title.normalizedISBN);
                mergedInventory.push(fullRecord);
            }
        });
        
        // Sort by ISBN
        mergedInventory.sort((a, b) => a.ISBN - b.ISBN);
        
        const jsonContent = JSON.stringify(mergedInventory, null, 2);
        const blob = new Blob([jsonContent], { type: 'application/json' });
        const url = window.URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'data.json'; // Simple filename without timestamp
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        const changesSummary = [];
        if (validNewTitles.length > 0) changesSummary.push(`${validNewTitles.length} titles added`);
        if (titlesToRemove.length > 0) changesSummary.push(`${titlesToRemove.length} titles removed`);
        
        showInventoryStatus(`📥 Updated repo downloaded as data.json with ${changesSummary.join(', ')}!`, 'success');
        
        secureLog('Repo update downloaded:', {
            added: validNewTitles.length,
            removed: titlesToRemove.length,
            total: mergedInventory.length
        });
        
    } catch (error) {
        console.error('Download error:', error);
        showInventoryStatus('Error creating updated repo file: ' + error.message, 'danger');
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
        
        showInventoryStatus('Current repo backup downloaded!', 'success');
        
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
}

// Initialize the application when DOM is ready
function initializeApp() {
    // Add event listeners only after DOM is loaded
    const excelFile = document.getElementById('excelFile');
    const inventoryFile = document.getElementById('inventoryFile');
    const removeFile = document.getElementById('removeFile');
    const removeIsbn = document.getElementById('removeIsbn');
    const isbnSearch = document.getElementById('isbnSearch');
    
    if (excelFile) {
        excelFile.addEventListener('change', handleFileSelect);
    }
    
    if (inventoryFile) {
        inventoryFile.addEventListener('change', function() {
            const processBtn = document.getElementById('processInventoryBtn');
            if (processBtn) {
                processBtn.disabled = !this.files[0];
            }
        });
    }
    
    if (removeFile) {
        removeFile.addEventListener('change', function() {
            const processBtn = document.getElementById('processBulkRemovalBtn');
            if (processBtn) {
                processBtn.disabled = !this.files[0];
            }
        });
    }
    
    if (removeIsbn) {
        removeIsbn.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                searchAndMarkForRemoval();
            }
        });
    }
    
    if (isbnSearch) {
        isbnSearch.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                searchISBN();
            }
        });
    }
    
    // Load initial data
    secureLog('Initializing application...');
    fetchData();
    
    // Start session monitoring if password protection is enabled
    if (INVENTORY_ACCESS.enabled) {
        // Check session every minute
        setInterval(checkInventorySession, 60000);
        secureLog('Session monitoring started');
    }
}

// Wait for DOM to be fully loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    // DOM is already loaded
    initializeApp();
}