// Global variables
let booksData = [];
let processedOrders = [];

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
        'application/csv'
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
        throw new Error('Invalid file type. Only Excel (.xlsx, .xls) and CSV (.csv) files are allowed');
    }
    
    // Check file extension as additional validation
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
        throw new Error('Invalid file extension. Only .xlsx, .xls, and .csv files are allowed');
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
        
        // Store original raw data globally for search functionality
        window.originalRawData = rawData;
        
        // Debug: Let's see the structure of the first few items
        console.log('First 3 raw data items:', rawData.slice(0, 3));
        console.log('Available fields in first item:', Object.keys(rawData[0] || {}));
        console.log('Sample ISBN values from raw data:', rawData.slice(0, 5).map(item => ({
            original: item.ISBN,
            type: typeof item.ISBN,
            keys: Object.keys(item).filter(key => key.toLowerCase().includes('isbn'))
        })));
        
        // Process and normalize the data with security checks
        booksData = rawData.map((item, index) => {
            // Debug first few items more thoroughly
            if (index < 3) {
                console.log(`Item ${index + 1} raw data:`, item);
                console.log(`Item ${index + 1} ISBN field:`, item.ISBN, 'Type:', typeof item.ISBN);
                console.log(`Item ${index + 1} all keys:`, Object.keys(item));
            }
            
            // Convert numeric ISBNs to strings, but be more careful about the field name
            let isbn = '';
            
            // Try different possible ISBN field names
            const isbnValue = item.ISBN || item.isbn || item['ISBN-13'] || item['isbn-13'] || item.ean || item.EAN;
            
            if (isbnValue !== undefined && isbnValue !== null) {
                isbn = String(isbnValue).trim();
                
                // Handle scientific notation that might come from Excel/JSON
                if (isbn.includes('e') || isbn.includes('E')) {
                    isbn = Number(isbn).toFixed(0);
                }
                
                // Only process if we have a valid-looking ISBN
                if (isbn && /^\d+$/.test(isbn) && isbn.length >= 10) {
                    // Normalize to 13 digits if it's 10-12 digits
                    if (isbn.length >= 10 && isbn.length <= 12) {
                        isbn = isbn.padStart(13, '0');
                    }
                }
            }
            
            // Debug the processed ISBN for first few items
            if (index < 3) {
                console.log(`Item ${index + 1} processed ISBN:`, isbn);
            }
            
            // Sanitize title and master order ID - handle both "Title" and "TITLE"
            const title = sanitizeText(item.Title || item.TITLE || item.title || 'No title available');
            const masterOrderId = sanitizeText(item['Master Order ID'] || item.masterOrderId || '');
            const status = sanitizeText(item.Status || item.status || 'POD Ready'); // Default to POD Ready if missing
            
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
                // Store original item for reference
                originalData: item,
                // Legacy compatibility
                code: isbn,
                description: title,
                setupdate: sanitizeText(item.setupdate || '')
            };
        });
        
        secureLog('Processed books data:', booksData);
        showStatus(`Successfully loaded ${booksData.length} books from repository`, 'success');
        updateRepositoryStats();
        
    } catch (error) {
        console.error('Error fetching data:', error);
        showStatus('Error loading book data. Please ensure data.json exists and is properly formatted.', 'danger');
        booksData = [];
    }
}

// Update repository statistics dashboard
function updateRepositoryStats() {
    // Count titles by status
    const podReadyTitles = booksData.filter(book => book.status === 'POD Ready').length;
    const mpiTitles = booksData.filter(book => book.status === 'MPI').length;
    
    // Better duplicate detection - exclude empty ISBNs
    const uniqueISBNs = new Set();
    const duplicateISBNs = new Set();
    let duplicatesCount = 0;
    
    booksData.forEach(book => {
        const isbn = book.isbn;
        // Only process non-empty ISBNs
        if (isbn && isbn.trim() !== '') {
            if (uniqueISBNs.has(isbn)) {
                // This ISBN has been seen before, it's a duplicate
                duplicateISBNs.add(isbn);
            } else {
                uniqueISBNs.add(isbn);
            }
        }
    });
    
    // Count total duplicate entries (excluding empty ISBNs)
    duplicateISBNs.forEach(duplicateISBN => {
        const count = booksData.filter(book => book.isbn === duplicateISBN).length;
        duplicatesCount += (count - 1); // Subtract 1 to count only the extra copies
    });
    
    // Debug information - temporarily enable to see what's happening
    console.log('Sample ISBNs from data:', booksData.slice(0, 5).map(b => b.isbn));
    console.log('Sample raw data:', booksData.slice(0, 2));
    console.log('Unique ISBNs (non-empty):', uniqueISBNs.size);
    console.log('Total entries:', booksData.length);
    console.log('Entries with ISBNs:', booksData.filter(b => b.isbn && b.isbn.trim() !== '').length);
    console.log('Duplicate ISBNs found:', duplicateISBNs.size);
    console.log('Total duplicate entries:', duplicatesCount);
    
    document.getElementById('currentTitlesCount').textContent = booksData.length;
    document.getElementById('podReadyCount').textContent = podReadyTitles;
    document.getElementById('mpiCount').textContent = mpiTitles;
    document.getElementById('duplicatesCount').textContent = duplicatesCount;
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
            
            // Look up book in repository
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
        
        showStatus(`Data loaded successfully! ${validCount}/${totalCount} items found in repository (${podReadyCount} POD Ready, ${mpiCount} MPI, ${notAvailableCount} Not Available).`, 'success');
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
        // Show only items not in repository
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
            cell.textContent = 'All items are available in repository!';
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
        
        // Updated status logic based on Status field
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
    
    console.log('Searching for:', rawInput);
    
    // Search directly in the original raw data since processed data has empty ISBNs
    if (window.originalRawData && window.originalRawData.length > 0) {
        
        // First try to search as ISBN
        if (validateISBN(rawInput)) {
            const normalizedSearchISBN = rawInput.replace(/\D/g, '');
            console.log('Normalized search ISBN:', normalizedSearchISBN);
            
            const rawFound = window.originalRawData.find(item => {
                // Try different possible ISBN field names and formats
                const possibleISBNs = [
                    String(item.ISBN || ''),
                    String(item.isbn || ''),
                    String(item['ISBN-13'] || ''),
                    String(item['isbn-13'] || '')
                ];
                
                return possibleISBNs.some(isbn => {
                    if (!isbn) return false;
                    const normalizedISBN = isbn.replace(/\D/g, '');
                    return normalizedISBN === normalizedSearchISBN || 
                           normalizedISBN.padStart(13, '0') === normalizedSearchISBN.padStart(13, '0');
                });
            });
            
            if (rawFound) {
                const displayISBN = String(rawFound.ISBN || rawFound.isbn || '').replace(/\D/g, '');
                foundBook = {
                    isbn: displayISBN.length >= 10 ? displayISBN.padStart(13, '0') : displayISBN,
                    title: sanitizeText(rawFound.Title || rawFound.TITLE || rawFound.title || 'No title available'),
                    masterOrderId: sanitizeText(rawFound['Master Order ID'] || rawFound.masterOrderId || ''),
                    status: sanitizeText(rawFound.Status || rawFound.status || 'POD Ready')
                };
                searchType = 'ISBN';
            }
            
            console.log('ISBN search result:', foundBook);
        }
        
        // If not found, try Master Order ID
        if (!foundBook) {
            const rawFound = window.originalRawData.find(item => {
                const masterOrderIds = [
                    item['Master Order ID'],
                    item.masterOrderId,
                    item.MasterOrderId,
                    item['Master Order Id']
                ];
                
                return masterOrderIds.some(id => 
                    id && String(id).toLowerCase() === rawInput.toLowerCase()
                );
            });
            
            if (rawFound) {
                const displayISBN = String(rawFound.ISBN || rawFound.isbn || '').replace(/\D/g, '');
                foundBook = {
                    isbn: displayISBN.length >= 10 ? displayISBN.padStart(13, '0') : displayISBN,
                    title: sanitizeText(rawFound.Title || rawFound.TITLE || rawFound.title || 'No title available'),
                    masterOrderId: sanitizeText(rawFound['Master Order ID'] || rawFound.masterOrderId || ''),
                    status: sanitizeText(rawFound.Status || rawFound.status || 'POD Ready')
                };
                searchType = 'Master Order ID';
            }
            
            console.log('Master Order ID search result:', foundBook);
        }
    } else {
        console.log('No original raw data available for search');
    }
    
    if (foundBook) {
        showISBNResult(`
            <strong>✓ Available</strong><br>
            <strong>ISBN:</strong> ${foundBook.isbn || 'Not available'}<br>
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
            This title is not available in the current repository.
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
            const dtlRow = [
                'DTL',
                orderRef,
                order.lineNumber,
                order.isbn,
                order.quantity.toString()
            ];
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

// Download Excel template function
function downloadTemplate() {
    try {
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
        
        showStatus('Excel template downloaded successfully!', 'success');
        
    } catch (error) {
        console.error('Template download error:', error);
        showStatus('Error downloading template. Please ensure order_template.xlsx exists on the server.', 'danger');
    }
}

// Download repository data function
function downloadRepositoryData() {
    if (booksData.length === 0) {
        showStatus('No repository data available to download', 'warning');
        return;
    }

    try {
        // Prepare data for Excel export with user-friendly column names
        const excelData = booksData.map(book => {
            // Get original data from the raw JSON structure
            return {
                'ISBN': book.isbn,
                'Master Order ID': book.masterOrderId,
                'Title': book.title,
                'Status': book.status,
                'Trim Height': book.trimHeight || '',
                'Trim Width': book.trimWidth || '',
                'Bind Style': book.bindStyle || '',
                'Extent': book.extent || '',
                'Paper Description': book.paperDesc || '',
                'Cover Spec Code': book.coverSpecCode1 || '',
                'Cover Spine': book.coverSpine || '',
                'Packing': book.packing || ''
            };
        });

        // Create Excel workbook
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(excelData);
        
        // Auto-size columns
        const maxWidths = {};
        Object.keys(excelData[0] || {}).forEach(key => {
            maxWidths[key] = Math.max(
                key.length,
                ...excelData.map(row => String(row[key] || '').length)
            );
        });
        
        ws['!cols'] = Object.keys(maxWidths).map(key => ({
            wch: Math.min(maxWidths[key] + 2, 50) // Cap at 50 characters width
        }));
        
        XLSX.utils.book_append_sheet(wb, ws, 'Repository Data');
        
        // Generate filename with timestamp
        const now = new Date();
        const filename = `repository_data_${now.getFullYear()}_${
            String(now.getMonth() + 1).padStart(2, '0')}_${
            String(now.getDate()).padStart(2, '0')}_${
            String(now.getHours()).padStart(2, '0')}_${
            String(now.getMinutes()).padStart(2, '0')}.xlsx`;
        
        // Download Excel file
        XLSX.writeFile(wb, filename);
        
        showStatus(`Repository data downloaded as Excel file: ${filename}`, 'success');
        
    } catch (error) {
        console.error('Excel download error:', error);
        
        // Fallback to CSV if Excel fails
        try {
            const csvData = booksData.map(book => ({
                'ISBN': book.isbn,
                'Master Order ID': book.masterOrderId,
                'Title': book.title,
                'Status': book.status,
                'Trim Height': book.trimHeight || '',
                'Trim Width': book.trimWidth || '',
                'Bind Style': book.bindStyle || '',
                'Extent': book.extent || '',
                'Paper Description': book.paperDesc || '',
                'Cover Spec Code': book.coverSpecCode1 || '',
                'Cover Spine': book.coverSpine || '',
                'Packing': book.packing || ''
            }));
            
            const csv = Papa.unparse(csvData);
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = window.URL.createObjectURL(blob);
            
            const a = document.createElement('a');
            a.href = url;
            const now = new Date();
            const filename = `repository_data_${now.getFullYear()}_${
                String(now.getMonth() + 1).padStart(2, '0')}_${
                String(now.getDate()).padStart(2, '0')}_${
                String(now.getHours()).padStart(2, '0')}_${
                String(now.getMinutes()).padStart(2, '0')}.csv`;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            showStatus(`Repository data downloaded as CSV file: ${filename} (Excel export failed)`, 'success');
            
        } catch (csvError) {
            console.error('CSV download error:', csvError);
            showStatus('Error downloading repository data: ' + csvError.message, 'danger');
        }
    }
}

// Initialize the application when DOM is ready
function initializeApp() {
    // Add event listeners only after DOM is loaded
    const excelFile = document.getElementById('excelFile');
    const isbnSearch = document.getElementById('isbnSearch');
    
    if (excelFile) {
        excelFile.addEventListener('change', handleFileSelect);
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
}

// Wait for DOM to be fully loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    // DOM is already loaded
    initializeApp();
}