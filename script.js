// Application State Management
const AppState = {
    booksData: [],
    processedOrders: [],
    originalRawData: null
};

// Configuration
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

const SECURITY_CONFIG = {
    maxFileSize: 10 * 1024 * 1024, // 10MB
    allowedMimeTypes: [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
        'application/csv'
    ],
    maxDescriptionLength: 200,
    enableDebugLogging: false
};

// Security and Validation Module
const SecurityModule = {
    sanitizeText(text) {
        if (typeof text !== 'string') return '';
        return text.replace(/[<>'"&]/g, '')
                   .substring(0, SECURITY_CONFIG.maxDescriptionLength)
                   .trim();
    },

    validateFile(file) {
        if (!file) {
            throw new Error('No file provided');
        }
        
        if (file.size > SECURITY_CONFIG.maxFileSize) {
            throw new Error(`File size exceeds limit of ${SECURITY_CONFIG.maxFileSize / (1024 * 1024)}MB`);
        }
        
        if (!SECURITY_CONFIG.allowedMimeTypes.includes(file.type)) {
            throw new Error('Invalid file type. Only Excel (.xlsx, .xls) and CSV (.csv) files are allowed');
        }
        
        const fileName = file.name.toLowerCase();
        if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
            throw new Error('Invalid file extension. Only .xlsx, .xls, and .csv files are allowed');
        }
        
        return true;
    },

    validateISBN(isbn) {
        if (typeof isbn !== 'string') return false;
        const cleanISBN = isbn.replace(/\D/g, '');
        return cleanISBN.length >= 10 && cleanISBN.length <= 13;
    },

    validateQuantity(qty) {
        const quantity = parseInt(qty);
        return !isNaN(quantity) && quantity > 0 && quantity <= 10000;
    },

    secureLog(message, data = null) {
        if (SECURITY_CONFIG.enableDebugLogging) {
            console.log(message, data);
        }
    }
};

// Data Processing Module
const DataProcessor = {
    normalizeISBN(isbn) {
        if (!isbn) return '';
        
        let normalized = String(isbn).trim();
        
        // Handle scientific notation
        if (normalized.includes('e') || normalized.includes('E')) {
            normalized = Number(normalized).toFixed(0);
        }
        
        // Remove non-digits and pad to 13 characters
        normalized = normalized.replace(/\D/g, '');
        if (normalized.length >= 10 && normalized.length <= 12) {
            normalized = normalized.padStart(13, '0');
        }
        
        return normalized;
    },

    extractISBN(item) {
        const possibleFields = ['ISBN', 'isbn', 'ISBN-13', 'isbn-13', 'ean', 'EAN'];
        for (const field of possibleFields) {
            if (item[field] !== undefined && item[field] !== null) {
                return this.normalizeISBN(item[field]);
            }
        }
        return '';
    },

    processBookData(rawData) {
        return rawData.map((item, index) => {
            const isbn = this.extractISBN(item);
            const title = SecurityModule.sanitizeText(item.Title || item.TITLE || item.title || 'No title available');
            const masterOrderId = SecurityModule.sanitizeText(item['Master Order ID'] || item.masterOrderId || '');
            const status = SecurityModule.sanitizeText(item.Status || item.status || 'POD Ready');
            const paperDesc = SecurityModule.sanitizeText(item['Paper Desc'] || item.paperDesc || item['Paper Description'] || '');
            
            return {
                isbn,
                title,
                masterOrderId,
                status,
                paperDesc,
                trimHeight: item['Trim Height'],
                trimWidth: item['Trim Width'],
                bindStyle: SecurityModule.sanitizeText(item['Bind Style'] || ''),
                extent: item.Extent,
                coverSpecCode1: SecurityModule.sanitizeText(item['Cover Spec Code 1'] || ''),
                coverSpine: item['Cover Spine'],
                packing: SecurityModule.sanitizeText(item.Packing || ''),
                originalData: item,
                // Legacy compatibility
                code: isbn,
                description: title,
                setupdate: SecurityModule.sanitizeText(item.setupdate || '')
            };
        });
    },

    createBooksMaps(booksData) {
        const booksMap = new Map();
        const masterOrderMap = new Map();
        
        booksData.forEach(item => {
            if (item.isbn) {
                booksMap.set(item.isbn, item);
            }
            if (item.masterOrderId) {
                masterOrderMap.set(item.masterOrderId.toLowerCase(), item);
            }
        });
        
        return { booksMap, masterOrderMap };
    }
};

// Search Module
const SearchModule = {
    findByISBN(searchISBN) {
        if (!AppState.originalRawData || !SecurityModule.validateISBN(searchISBN)) {
            return null;
        }
        
        const normalizedSearchISBN = DataProcessor.normalizeISBN(searchISBN);
        
        return AppState.originalRawData.find(item => {
            const itemISBN = DataProcessor.extractISBN(item);
            return itemISBN === normalizedSearchISBN;
        });
    },

    findByMasterOrderId(searchId) {
        if (!AppState.originalRawData) return null;
        
        const possibleFields = ['Master Order ID', 'masterOrderId', 'MasterOrderId', 'Master Order Id'];
        
        return AppState.originalRawData.find(item => {
            return possibleFields.some(field => {
                const value = item[field];
                return value && String(value).toLowerCase() === searchId.toLowerCase();
            });
        });
    },

    createBookFromRaw(rawItem) {
        if (!rawItem) return null;
        
        const isbn = DataProcessor.extractISBN(rawItem);
        return {
            isbn: isbn,
            title: SecurityModule.sanitizeText(rawItem.Title || rawItem.TITLE || rawItem.title || 'No title available'),
            masterOrderId: SecurityModule.sanitizeText(rawItem['Master Order ID'] || rawItem.masterOrderId || ''),
            status: SecurityModule.sanitizeText(rawItem.Status || rawItem.status || 'POD Ready'),
            paperDesc: SecurityModule.sanitizeText(rawItem['Paper Desc'] || rawItem.paperDesc || rawItem['Paper Description'] || 'Not specified')
        };
    }
};

// UI Module
const UIModule = {
    showStatus(message, type) {
        const statusDiv = document.getElementById('status');
        statusDiv.className = `alert alert-${type}`;
        statusDiv.textContent = message;
        statusDiv.style.display = 'block';
        
        if (type === 'success' || type === 'info') {
            setTimeout(() => {
                statusDiv.style.display = 'none';
            }, 5000);
        }
    },

    enableButtons(enabled) {
        const buttonIds = ['clearBtn', 'downloadBtn', 'deleteSelectedBtn', 'copyBtn'];
        buttonIds.forEach(id => {
            const button = document.getElementById(id);
            if (button) button.disabled = !enabled;
        });
        
        const selectAll = document.getElementById('selectAll');
        if (selectAll) selectAll.checked = false;
    },

    updateRepositoryStats() {
        const podReadyCount = AppState.booksData.filter(book => book.status === 'POD Ready').length;
        const mpiCount = AppState.booksData.filter(book => book.status === 'MPI').length;
        
        // Improved duplicate detection
        const isbnCounts = new Map();
        AppState.booksData.forEach(book => {
            if (book.isbn && book.isbn.trim() !== '') {
                isbnCounts.set(book.isbn, (isbnCounts.get(book.isbn) || 0) + 1);
            }
        });
        
        const duplicatesCount = Array.from(isbnCounts.values())
            .filter(count => count > 1)
            .reduce((total, count) => total + (count - 1), 0);
        
        document.getElementById('currentTitlesCount').textContent = AppState.booksData.length;
        document.getElementById('podReadyCount').textContent = podReadyCount;
        document.getElementById('mpiCount').textContent = mpiCount;
        document.getElementById('duplicatesCount').textContent = duplicatesCount;
    },

    getStatusBadgeClass(status) {
        switch (status) {
            case 'POD Ready':
                return 'bg-success';
            case 'MPI':
                return 'bg-warning text-dark';
            default:
                return 'bg-info';
        }
    },

    showISBNResult(content, type) {
        const resultDiv = document.getElementById('isbnResult');
        const alertDiv = document.getElementById('isbnResultAlert');
        const contentDiv = document.getElementById('isbnResultContent');
        
        alertDiv.className = `alert alert-${type}`;
        contentDiv.innerHTML = content;
        resultDiv.style.display = 'block';
        
        if (type === 'success' || type === 'info') {
            setTimeout(() => {
                resultDiv.style.display = 'none';
            }, 10000);
        }
    },

    updatePreviewTable() {
        const tbody = document.getElementById('previewBody');
        tbody.innerHTML = '';

        if (AppState.processedOrders.length === 0) {
            const row = document.createElement('tr');
            const cell = document.createElement('td');
            cell.colSpan = 7;
            cell.className = 'text-center text-muted';
            cell.innerHTML = '<i class="fas fa-inbox"></i> No data loaded';
            row.appendChild(cell);
            tbody.appendChild(row);
            return;
        }

        const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
        const showNotAvailable = document.getElementById('showNotAvailable')?.checked || false;
        
        let filteredOrders = AppState.processedOrders;

        if (printAsMiscellaneous) {
            filteredOrders = AppState.processedOrders.filter(order => order.available && order.status === 'MPI');
        } else if (showNotAvailable) {
            filteredOrders = AppState.processedOrders.filter(order => !order.available);
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

        filteredOrders.forEach(order => {
            const originalIndex = AppState.processedOrders.indexOf(order);
            const tr = document.createElement('tr');
            
            // Checkbox cell
            const checkboxCell = document.createElement('td');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.className = 'row-checkbox';
            checkbox.dataset.index = originalIndex;
            checkboxCell.appendChild(checkbox);
            
            // Action cell
            const actionCell = document.createElement('td');
            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'btn btn-danger btn-sm';
            deleteBtn.onclick = () => OrderProcessor.deleteRow(originalIndex);
            deleteBtn.innerHTML = '<i class="fas fa-trash"></i>';
            actionCell.appendChild(deleteBtn);
            
            // Data cells
            const lineCell = document.createElement('td');
            lineCell.textContent = order.lineNumber;
            
            const isbnCell = document.createElement('td');
            isbnCell.textContent = order.isbn;
            
            const descCell = document.createElement('td');
            descCell.textContent = order.description;
            
            const qtyCell = document.createElement('td');
            qtyCell.textContent = order.quantity;
            
            // Status cell with badge
            const statusCell = document.createElement('td');
            const badge = document.createElement('span');
            
            if (!order.available) {
                badge.className = 'badge bg-danger';
                badge.textContent = 'Not Available';
            } else {
                badge.className = `badge ${this.getStatusBadgeClass(order.status)}`;
                badge.textContent = order.status;
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
};

// File Processing Module
const FileProcessor = {
    async processFile(file, orderRef) {
        SecurityModule.validateFile(file);
        
        const arrayBuffer = await file.arrayBuffer();
        let excelData = [];
        
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
            const workbook = XLSX.read(arrayBuffer);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(worksheet);
        }

        return this.processOrderData(excelData, orderRef);
    },

    processOrderData(excelData, orderRef) {
        const { booksMap } = DataProcessor.createBooksMaps(AppState.booksData);
        
        return excelData.map((row, index) => {
            // Extract and normalize ISBN
            let isbn = String(row.ISBN || row.isbn || '').trim();
            
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            if (!SecurityModule.validateISBN(isbn)) {
                isbn = '';
            } else {
                isbn = DataProcessor.normalizeISBN(isbn);
            }
            
            // Extract and validate quantity
            const rawQuantity = row.Qty || row.qty || row.Quantity || row.quantity || 0;
            const quantity = SecurityModule.validateQuantity(rawQuantity) ? parseInt(rawQuantity) : 0;
            
            // Look up in repository
            const stockItem = booksMap.get(isbn);
            
            return {
                lineNumber: String(index + 1).padStart(3, '0'),
                orderRef: SecurityModule.sanitizeText(orderRef),
                isbn,
                description: stockItem?.title || 'Not Found',
                quantity,
                available: !!stockItem,
                status: stockItem?.status || 'Not Available',
                masterOrderId: stockItem?.masterOrderId || ''
            };
        });
    }
};

// Order Processing Module
const OrderProcessor = {
    deleteRow(index) {
        AppState.processedOrders.splice(index, 1);
        AppState.processedOrders = AppState.processedOrders.map((order, idx) => ({
            ...order,
            lineNumber: String(idx + 1).padStart(3, '0')
        }));
        UIModule.updatePreviewTable();
        UIModule.showStatus(`Row ${index + 1} deleted`, 'info');
        UIModule.enableButtons(AppState.processedOrders.length > 0);
    },

    clearAll() {
        AppState.processedOrders = [];
        UIModule.updatePreviewTable();
        document.getElementById('excelFile').value = '';
        document.getElementById('orderRef').value = '';
        UIModule.showStatus('All data cleared', 'info');
        UIModule.enableButtons(false);
    },

    deleteSelected() {
        const checkboxes = document.querySelectorAll('.row-checkbox:checked');
        const indices = Array.from(checkboxes)
            .map(checkbox => parseInt(checkbox.dataset.index))
            .sort((a, b) => b - a);

        indices.forEach(index => {
            AppState.processedOrders.splice(index, 1);
        });

        AppState.processedOrders = AppState.processedOrders.map((order, idx) => ({
            ...order,
            lineNumber: String(idx + 1).padStart(3, '0')
        }));

        UIModule.updatePreviewTable();
        UIModule.showStatus(`${indices.length} rows deleted`, 'info');
        UIModule.enableButtons(AppState.processedOrders.length > 0);
    }
};

// Export Module
const ExportModule = {
    formatDate() {
        const now = new Date();
        return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    },

    downloadCsv() {
        if (AppState.processedOrders.length === 0) {
            UIModule.showStatus('No data to download', 'warning');
            return;
        }

        try {
            const formattedDate = this.formatDate();
            const orderRef = SecurityModule.sanitizeText(document.getElementById('orderRef').value);
        
            const csvRow = customerConfig.csvStructure.map(field => {
                switch (field) {
                    case 'HDR': return 'HDR';
                    case 'orderNumber': return orderRef;
                    case 'date': return formattedDate;
                    case 'type': return customerConfig.type || '';
                    case 'companyName': return customerConfig.name;
                    case 'phone': return customerConfig.phone;
                    default: return customerConfig.address[field] || '';
                }
            });
        
            const finalCsvRow = [...csvRow, ''];
            const csvContent = [finalCsvRow];
        
            AppState.processedOrders.forEach(order => {
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
        
            UIModule.showStatus('CSV downloaded successfully!', 'success');
        } catch (error) {
            console.error('Download error:', error);
            UIModule.showStatus('Error creating CSV file: ' + error.message, 'danger');
        }
    },

    copyTableToClipboard() {
        if (AppState.processedOrders.length === 0) return;

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

        AppState.processedOrders.forEach(order => {
            const safeDescription = order.description.replace(/[<>&"']/g, function(match) {
                const escapeMap = {
                    '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&#39;'
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
            UIModule.showStatus('Table copied to clipboard', 'success');
        }).catch(err => {
            console.error('Failed to copy:', err);
            UIModule.showStatus('Failed to copy table', 'danger');
        });
    },

    downloadRepositoryData() {
        if (AppState.booksData.length === 0) {
            UIModule.showStatus('No repository data available to download', 'warning');
            return;
        }

        try {
            const excelData = AppState.booksData.map(book => ({
                'ISBN': book.isbn,
                'Master Order ID': book.masterOrderId,
                'Title': book.title,
                'Status': book.status,
                'Paper Description': book.paperDesc,
                'Trim Height': book.trimHeight || '',
                'Trim Width': book.trimWidth || '',
                'Bind Style': book.bindStyle || '',
                'Extent': book.extent || '',
                'Cover Spec Code': book.coverSpecCode1 || '',
                'Cover Spine': book.coverSpine || '',
                'Packing': book.packing || ''
            }));

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
                wch: Math.min(maxWidths[key] + 2, 50)
            }));
            
            XLSX.utils.book_append_sheet(wb, ws, 'Repository Data');
            
            const now = new Date();
            const filename = `repository_data_${now.getFullYear()}_${
                String(now.getMonth() + 1).padStart(2, '0')}_${
                String(now.getDate()).padStart(2, '0')}_${
                String(now.getHours()).padStart(2, '0')}_${
                String(now.getMinutes()).padStart(2, '0')}.xlsx`;
            
            XLSX.writeFile(wb, filename);
            UIModule.showStatus(`Repository data downloaded as Excel file: ${filename}`, 'success');
            
        } catch (error) {
            console.error('Excel download error:', error);
            this.downloadRepositoryDataAsCsv(); // Fallback to CSV
        }
    },

    downloadRepositoryDataAsCsv() {
        try {
            const csvData = AppState.booksData.map(book => ({
                'ISBN': book.isbn,
                'Master Order ID': book.masterOrderId,
                'Title': book.title,
                'Status': book.status,
                'Paper Description': book.paperDesc,
                'Trim Height': book.trimHeight || '',
                'Trim Width': book.trimWidth || '',
                'Bind Style': book.bindStyle || '',
                'Extent': book.extent || '',
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
            
            UIModule.showStatus(`Repository data downloaded as CSV file: ${filename}`, 'success');
            
        } catch (csvError) {
            console.error('CSV download error:', csvError);
            UIModule.showStatus('Error downloading repository data: ' + csvError.message, 'danger');
        }
    },

    downloadTemplate() {
        try {
            const link = document.createElement('a');
            link.href = 'order_template.xlsx';
            link.download = 'order_template.xlsx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            UIModule.showStatus('Excel template downloaded successfully!', 'success');
        } catch (error) {
            console.error('Template download error:', error);
            UIModule.showStatus('Error downloading template. Please ensure order_template.xlsx exists on the server.', 'danger');
        }
    }
};

// Main Application Logic
async function fetchData() {
    try {
        SecurityModule.secureLog('Attempting to fetch data.json...');
        const response = await fetch('data.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const rawData = await response.json();
        
        // Store original data for search functionality
        AppState.originalRawData = rawData;
        
        // Process data with improved structure
        AppState.booksData = DataProcessor.processBookData(rawData);
        
        UIModule.showStatus(`Successfully loaded ${AppState.booksData.length} books from repository`, 'success');
        UIModule.updateRepositoryStats();
        
    } catch (error) {
        console.error('Error fetching data:', error);
        UIModule.showStatus('Error loading book data. Please ensure data.json exists and is properly formatted.', 'danger');
        AppState.booksData = [];
    }
}

async function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    const orderRef = SecurityModule.sanitizeText(document.getElementById('orderRef').value);
    const orderRefWarning = document.getElementById('orderRefWarning');
    
    if (!orderRef) {
        orderRefWarning.style.display = 'block';
        UIModule.showStatus('Please enter an order reference before uploading a file', 'warning');
        document.getElementById('excelFile').value = '';
        return;
    }
    
    orderRefWarning.style.display = 'none';

    try {
        UIModule.showStatus('Processing file...', 'info');
        
        AppState.processedOrders = await FileProcessor.processFile(file, orderRef);
        UIModule.updatePreviewTable();
        
        const validCount = AppState.processedOrders.filter(order => order.available).length;
        const totalCount = AppState.processedOrders.length;
        const podReadyCount = AppState.processedOrders.filter(order => order.available && order.status === 'POD Ready').length;
        const mpiCount = AppState.processedOrders.filter(order => order.available && order.status === 'MPI').length;
        const notAvailableCount = AppState.processedOrders.filter(order => !order.available).length;
        
        UIModule.showStatus(`Data loaded successfully! ${validCount}/${totalCount} items found in repository (${podReadyCount} POD Ready, ${mpiCount} MPI, ${notAvailableCount} Not Available).`, 'success');
        UIModule.enableButtons(true);
        
    } catch (error) {
        console.error('Processing error:', error);
        UIModule.showStatus('Error processing file: ' + error.message, 'danger');
        UIModule.enableButtons(false);
        document.getElementById('excelFile').value = '';
    }
}

function searchISBN() {
    const isbnInput = document.getElementById('isbnSearch');
    const rawInput = SecurityModule.sanitizeText(isbnInput.value);
    
    if (!rawInput) {
        UIModule.showISBNResult('Please enter an ISBN or Master Order ID to search', 'warning');
        return;
    }
    
    let foundBook = null;
    let searchType = '';
    
    console.log('Searching for:', rawInput);
    
    // Try ISBN search first
    if (SecurityModule.validateISBN(rawInput)) {
        const rawFound = SearchModule.findByISBN(rawInput);
        if (rawFound) {
            foundBook = SearchModule.createBookFromRaw(rawFound);
            searchType = 'ISBN';
        }
    }
    
    // If not found, try Master Order ID
    if (!foundBook) {
        const rawFound = SearchModule.findByMasterOrderId(rawInput);
        if (rawFound) {
            foundBook = SearchModule.createBookFromRaw(rawFound);
            searchType = 'Master Order ID';
        }
    }
    
    if (foundBook) {
        UIModule.showISBNResult(`
            <strong>✓ Available</strong><br>
            <strong>ISBN:</strong> ${foundBook.isbn || 'Not available'}<br>
            <strong>Title:</strong> ${foundBook.title}<br>
            <strong>Master Order ID:</strong> ${foundBook.masterOrderId || 'N/A'}<br>
            <strong>Paper:</strong> ${foundBook.paperDesc}<br>
            <strong>Status:</strong> <span class="badge ${UIModule.getStatusBadgeClass(foundBook.status)}">${foundBook.status}</span><br>
            <small class="text-muted">Found by ${searchType}</small>
        `, 'success');
    } else {
        const searchTypeMsg = SecurityModule.validateISBN(rawInput) ? 'ISBN' : 'Master Order ID';
        UIModule.showISBNResult(`
            <strong>✗ Not Available</strong><br>
            <strong>${searchTypeMsg}:</strong> ${rawInput}<br>
            This title is not available in the current repository.
        `, 'danger');
    }
}

function toggleTableFilter() {
    UIModule.updatePreviewTable();
    document.getElementById('selectAll').checked = false;
    
    const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
    const showNotAvailable = document.getElementById('showNotAvailable')?.checked || false;
    
    // Ensure only one filter is active at a time
    if (printAsMiscellaneous && showNotAvailable) {
        if (event.target.id === 'showOnlyUnavailable') {
            document.getElementById('showNotAvailable').checked = false;
        } else {
            document.getElementById('showOnlyUnavailable').checked = false;
        }
        setTimeout(toggleTableFilter, 10);
        return;
    }
    
    const mpiCount = AppState.processedOrders.filter(order => order.available && order.status === 'MPI').length;
    const notAvailableCount = AppState.processedOrders.filter(order => !order.available).length;
    
    if (printAsMiscellaneous) {
        UIModule.showStatus(`Showing ${mpiCount} MPI items`, 'info');
    } else if (showNotAvailable) {
        UIModule.showStatus(`Showing ${notAvailableCount} not available items`, 'info');
    } else {
        UIModule.showStatus(`Showing all ${AppState.processedOrders.length} items`, 'info');
    }
}

function toggleAllCheckboxes() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const selectAllCheckbox = document.getElementById('selectAll');
    checkboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
    });
}

// Global function exports for HTML onclick handlers
window.clearAll = OrderProcessor.clearAll;
window.downloadCsv = ExportModule.downloadCsv;
window.downloadRepositoryData = ExportModule.downloadRepositoryData;
window.downloadTemplate = ExportModule.downloadTemplate;
window.searchISBN = searchISBN;
window.deleteSelected = OrderProcessor.deleteSelected;
window.copyTableToClipboard = ExportModule.copyTableToClipboard;
window.toggleTableFilter = toggleTableFilter;
window.toggleAllCheckboxes = toggleAllCheckboxes;

// Initialize the application
function initializeApp() {
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
    
    SecurityModule.secureLog('Initializing application...');
    fetchData();
}

// Wait for DOM to be fully loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}