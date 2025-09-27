// Application State Management
const AppState = {
    booksData: [],
    processedOrders: [],
    originalRawData: null,
    booksMap: new Map(),
    masterOrderMap: new Map(),
    sortedByPaperType: false
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

    sanitizeXML(text) {
        if (typeof text !== 'string') return '';
        return text.replace(/&/g, '&amp;')
                   .replace(/</g, '&lt;')
                   .replace(/>/g, '&gt;')
                   .replace(/"/g, '&quot;')
                   .replace(/'/g, '&#39;');
    },

    validateFile(file) {
        if (!file) throw new Error('No file provided');
        
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
    }
};

// Data Processing Module (Optimized)
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
        const processedData = rawData.map(item => {
            const isbn = this.extractISBN(item);
            const title = SecurityModule.sanitizeText(item.Title || item.title || item.TITLE || 'No title available');
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
                coverSpec: SecurityModule.sanitizeText(item['Cover Spec'] || ''),
                coverSpine: item['Cover Spine'],
                packing: SecurityModule.sanitizeText(item.Packing || ''),
                originalData: item,
                // Legacy compatibility
                code: isbn,
                description: title,
                setupdate: SecurityModule.sanitizeText(item.setupdate || '')
            };
        });

        // Create optimized lookup maps
        this.createBooksMaps(processedData);
        return processedData;
    },

    createBooksMaps(booksData) {
        AppState.booksMap.clear();
        AppState.masterOrderMap.clear();
        
        booksData.forEach(item => {
            if (item.isbn) {
                AppState.booksMap.set(item.isbn, item);
            }
            if (item.masterOrderId) {
                AppState.masterOrderMap.set(item.masterOrderId.toLowerCase(), item);
            }
        });
    }
};

// Optimized Search Module
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
            title: SecurityModule.sanitizeText(rawItem.Title || rawItem.title || rawItem.TITLE || 'No title available'),
            masterOrderId: SecurityModule.sanitizeText(rawItem['Master Order ID'] || rawItem.masterOrderId || ''),
            status: SecurityModule.sanitizeText(rawItem.Status || rawItem.status || 'POD Ready'),
            paperDesc: SecurityModule.sanitizeText(rawItem['Paper Desc'] || rawItem.paperDesc || rawItem['Paper Description'] || 'Not specified')
        };
    }
};

// UI Module (Optimized)
const UIModule = {
    showStatus(message, type, persist = false) {
        const statusDiv = document.getElementById('status');
        statusDiv.className = `alert alert-${type}`;
        statusDiv.textContent = message;
        statusDiv.style.display = 'block';
        
        // Don't auto-hide if persist is true or if it's an import summary
        const isImportSummary = message.includes('Data loaded successfully');
        
        if (!persist && !isImportSummary) {
            if (type === 'success' || type === 'info') {
                setTimeout(() => {
                    statusDiv.style.display = 'none';
                }, 15000);
            }
        }
    },

    enableButtons(enabled) {
        const buttonIds = ['clearBtn', 'downloadBtn', 'deleteSelectedBtn', 'sortPaperBtn'];
        buttonIds.forEach(id => {
            const button = document.getElementById(id);
            if (button) button.disabled = !enabled;
        });
        
        const selectAll = document.getElementById('selectAll');
        if (selectAll) selectAll.checked = false;
    },

    enableBatchXML(enabled) {
        const csvUpload = document.getElementById('csvUpload');
        const batchBtn = document.getElementById('batchXMLBtn');
        if (csvUpload) csvUpload.disabled = !enabled;
        if (batchBtn) batchBtn.disabled = !enabled;
    },

    updateRepositoryStats() {
        const stats = AppState.booksData.reduce((acc, book) => {
            if (book.status === 'POD Ready') acc.podReady++;
            else if (book.status === 'MPI') acc.mpi++;
            return acc;
        }, { podReady: 0, mpi: 0 });
        
        // Optimized duplicate detection using Map
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
        document.getElementById('podReadyCount').textContent = stats.podReady;
        document.getElementById('mpiCount').textContent = stats.mpi;
        document.getElementById('duplicatesCount').textContent = duplicatesCount;
    },

    getStatusBadgeClass(status) {
        switch (status) {
            case 'POD Ready': return 'bg-success';
            case 'MPI': return 'bg-warning text-dark';
            default: return 'bg-info';
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
                const isbnSearchField = document.getElementById('isbnSearch');
                if (isbnSearchField) {
                    isbnSearchField.value = '';
                }
            }, 10000);
        }
    },

    updatePreviewTable() {
        const tbody = document.getElementById('previewBody');
        tbody.innerHTML = '';

        if (AppState.processedOrders.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted"><i class="fas fa-inbox"></i> No data loaded</td></tr>';
            return;
        }

        const filteredOrders = this.getFilteredOrders();

        if (filteredOrders.length === 0) {
            const message = this.getFilterMessage();
            tbody.innerHTML = `<tr><td colspan="8" class="text-center text-success">${message}</td></tr>`;
            return;
        }

        const fragment = document.createDocumentFragment();
        
        // Check if we should show paper type grouping headers
        const sortedByPaper = AppState.sortedByPaperType || false;
        let lastPaperType = null;
        let paperTypeQuantity = 0;
        
        filteredOrders.forEach((order, idx) => {
            const originalIndex = AppState.processedOrders.indexOf(order);
            
            // Add paper type grouping header if sorted by paper type
            if (sortedByPaper && order.paperDesc !== lastPaperType) {
                // Calculate total quantity for this paper type
                paperTypeQuantity = filteredOrders
                    .filter(o => o.paperDesc === order.paperDesc)
                    .reduce((sum, o) => sum + o.quantity, 0);
                
                const headerRow = document.createElement('tr');
                headerRow.className = 'table-secondary';
                headerRow.innerHTML = `
                    <td colspan="8" class="fw-bold">
                        <i class="fas fa-layer-group"></i> Paper Type: ${order.paperDesc || 'Not specified'} 
                        <span class="badge bg-primary ms-2">Total Qty: ${paperTypeQuantity}</span>
                    </td>
                `;
                fragment.appendChild(headerRow);
                lastPaperType = order.paperDesc;
            }
            
            const tr = this.createTableRow(order, originalIndex);
            fragment.appendChild(tr);
        });
        
        tbody.appendChild(fragment);
    },

    getFilteredOrders() {
        const showOnlyPodReady = document.getElementById('showOnlyPodReady')?.checked || false;
        const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
        const showNotAvailable = document.getElementById('showNotAvailable')?.checked || false;
        const paperTypeFilter = document.getElementById('paperTypeFilter')?.value || '';
        
        let filtered = AppState.processedOrders;
        
        // Apply status filters
        if (showOnlyPodReady) {
            filtered = filtered.filter(order => order.available && order.status === 'POD Ready');
        } else if (printAsMiscellaneous) {
            filtered = filtered.filter(order => order.available && order.status === 'MPI');
        } else if (showNotAvailable) {
            filtered = filtered.filter(order => !order.available);
        }
        
        // Apply paper type filter
        if (paperTypeFilter) {
            filtered = filtered.filter(order => order.paperDesc === paperTypeFilter);
        }
        
        return filtered;
    },

    getFilterMessage() {
        const showOnlyPodReady = document.getElementById('showOnlyPodReady')?.checked || false;
        const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
        const showNotAvailable = document.getElementById('showNotAvailable')?.checked || false;
        
        if (showOnlyPodReady) return 'No POD Ready items found!';
        if (printAsMiscellaneous) return 'No MPI items found!';
        if (showNotAvailable) return 'All items are available in repository!';
        return 'No data to display';
    },

    createTableRow(order, originalIndex) {
        const tr = document.createElement('tr');
        
        // Build row HTML efficiently - Status before Paper Type
        tr.innerHTML = `
            <td><input type="checkbox" class="row-checkbox" data-index="${originalIndex}"></td>
            <td><button class="btn btn-danger btn-sm" onclick="OrderProcessor.deleteRow(${originalIndex})"><i class="fas fa-trash"></i></button></td>
            <td>${order.lineNumber}</td>
            <td>${order.isbn}</td>
            <td>${order.description}</td>
            <td>${order.quantity}</td>
            <td><span class="badge ${this.getStatusBadgeClass(order.status)}">${order.available ? order.status : 'Not Available'}</span></td>
            <td><small>${order.paperDesc || 'Not specified'}</small></td>
        `;
        
        return tr;
    }
};

// File Processing Module (Optimized)
const FileProcessor = {
    async processFile(file, orderRef) {
        SecurityModule.validateFile(file);
        
        const arrayBuffer = await file.arrayBuffer();
        const excelData = await this.parseFileData(file, arrayBuffer);
        return this.processOrderData(excelData, orderRef);
    },

    async parseFileData(file, arrayBuffer) {
        if (file.name.toLowerCase().endsWith('.csv')) {
            const text = new TextDecoder().decode(arrayBuffer);
            const parsed = Papa.parse(text, {
                header: true,
                dynamicTyping: true,
                skipEmptyLines: true,
                delimitersToGuess: [',', '\t', '|', ';']
            });
            return parsed.data;
        } else {
            const workbook = XLSX.read(arrayBuffer);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            return XLSX.utils.sheet_to_json(worksheet);
        }
    },

    processOrderData(excelData, orderRef) {
        const orderItems = excelData.map((row, index) => {
            let isbn = '';
            let stockItem = null;
            let lookupMethod = '';
            
            // Try ISBN lookup first (original method)
            const rawISBN = String(row.ISBN || row.isbn || '').trim();
            if (rawISBN) {
                let processedISBN = rawISBN;
                
                if (processedISBN.includes('e') || processedISBN.includes('E')) {
                    processedISBN = Number(processedISBN).toFixed(0);
                }
                
                if (SecurityModule.validateISBN(processedISBN)) {
                    isbn = DataProcessor.normalizeISBN(processedISBN);
                    stockItem = AppState.booksMap.get(isbn);
                    if (stockItem) {
                        lookupMethod = 'ISBN';
                    }
                }
            }
            
            // If ISBN lookup failed, try Master Order ID lookup (new method)
            if (!stockItem) {
                const masterOrderId = String(row.Master || row.master || row['Master Order ID'] || '').trim();
                if (masterOrderId) {
                    stockItem = AppState.masterOrderMap.get(masterOrderId.toLowerCase());
                    if (stockItem) {
                        isbn = stockItem.isbn;
                        lookupMethod = 'Master Order ID';
                    }
                }
            }
            
            // Extract quantity - support multiple column names
            const rawQuantity = row.Qty || row.qty || row.Quantity || row.quantity || row.Rem || row.rem || 0;
            const quantity = SecurityModule.validateQuantity(rawQuantity) ? parseInt(rawQuantity) : 0;
            
            // Extract optional date field
            const orderDate = SecurityModule.sanitizeText(row.Date || row.date || '');
            
            return {
                originalIndex: index,
                orderRef: SecurityModule.sanitizeText(orderRef),
                isbn: isbn || '',
                description: stockItem?.title || row.Title || row.title || 'Not Found',
                quantity,
                available: !!stockItem,
                status: stockItem?.status || 'Not Available',
                masterOrderId: stockItem?.masterOrderId || '',
                orderDate: orderDate,
                lookupMethod: lookupMethod || 'Not Found',
                paperDesc: stockItem?.paperDesc || 'Not specified'
            };
        });
        
        // Consolidate duplicate ISBNs by summing quantities
        const consolidatedMap = new Map();
        
        orderItems.forEach(item => {
            const key = item.isbn || `NOT_FOUND_${item.originalIndex}`;
            
            if (consolidatedMap.has(key)) {
                const existing = consolidatedMap.get(key);
                existing.quantity += item.quantity;
                existing.consolidatedCount = (existing.consolidatedCount || 1) + 1;
            } else {
                consolidatedMap.set(key, { ...item, consolidatedCount: 1 });
            }
        });
        
        // Convert back to array and add line numbers
        const consolidatedOrders = Array.from(consolidatedMap.values()).map((order, idx) => ({
            ...order,
            lineNumber: String(idx + 1).padStart(3, '0')
        }));
        
        return consolidatedOrders;
    }
};

// Order Processing Module (Optimized)
const OrderProcessor = {
    deleteRow(index) {
        AppState.processedOrders.splice(index, 1);
        AppState.processedOrders = AppState.processedOrders.map((order, idx) => ({
            ...order,
            lineNumber: String(idx + 1).padStart(3, '0')
        }));
        
        // Repopulate paper type filter after deletion
        populatePaperTypeFilter();
        
        UIModule.updatePreviewTable();
        UIModule.showStatus(`Row ${index + 1} deleted`, 'info');
        UIModule.enableButtons(AppState.processedOrders.length > 0);
    },

    clearAll() {
        AppState.processedOrders = [];
        AppState.sortedByPaperType = false;
        
        // Clear paper type filter
        const paperTypeFilter = document.getElementById('paperTypeFilter');
        if (paperTypeFilter) {
            paperTypeFilter.innerHTML = '<option value="">All Paper Types</option>';
        }
        
        UIModule.updatePreviewTable();
        document.getElementById('excelFile').value = '';
        document.getElementById('orderRef').value = '';
        document.getElementById('csvUpload').value = '';
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

        // Repopulate paper type filter after deletion
        populatePaperTypeFilter();

        UIModule.updatePreviewTable();
        UIModule.showStatus(`${indices.length} rows deleted`, 'info');
        UIModule.enableButtons(AppState.processedOrders.length > 0);
    }
};

// Enhanced XML Generation Module with Batch Processing
const XMLModule = {
    generateXMLForBook(book) {
        if (!book) {
            throw new Error('No book data provided for XML generation');
        }

        const isbn = SecurityModule.sanitizeXML(String(book.ISBN || book.isbn || ''));
        const title = SecurityModule.sanitizeXML(book.TITLE || book.Title || book.title || 'LIVING (00)');
        const trimHeight = book['Trim Height'] || 216;
        const trimWidth = book['Trim Width'] || 135;
        const spineSize = book['Cover Spine'] || book['Spine Size'] || 17;
        const paperType = SecurityModule.sanitizeXML(book['Paper Desc'] || book['Paper Description'] || 'Holmen Cream 65 gsm');
        const bindingStyle = SecurityModule.sanitizeXML(book['Bind Style'] || book['Binding Style'] || 'Limp');
        const pageExtent = book.Extent || book['Page Extent'] || 240;

        return `<?xml version="1.0" encoding="UTF-8"?>
<book>
    <basic_info>
        <isbn>${isbn}</isbn>
        <title>${title}</title>
    </basic_info>
    <specifications>
        <dimensions>
            <trim_height>${trimHeight}</trim_height>
            <trim_width>${trimWidth}</trim_width>
            <spine_size>${spineSize}</spine_size>
        </dimensions>
        <materials>
            <paper_type>${paperType}</paper_type>
            <binding_style>${bindingStyle}</binding_style>
            <lamination>No Lamination</lamination>
        </materials>
        <page_extent>${pageExtent}</page_extent>
    </specifications>
</book>`;
    },

    downloadXMLFiles(book) {
        if (!book || (!book.isbn && !book.ISBN)) {
            UIModule.showStatus('No valid book data available for XML download', 'warning');
            return;
        }

        try {
            const xmlContent = this.generateXMLForBook(book);
            const isbn = book.ISBN || book.isbn;
            
            const downloadXML = (content, filename) => {
                const blob = new Blob([content], { type: 'application/xml' });
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = filename;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                window.URL.revokeObjectURL(url);
            };
            
            downloadXML(xmlContent, `${isbn}_c.xml`);
            
            setTimeout(() => {
                downloadXML(xmlContent, `${isbn}_t.xml`);
                UIModule.showStatus(`XML files downloaded: ${isbn}_c.xml and ${isbn}_t.xml`, 'success');
            }, 100);
            
        } catch (error) {
            console.error('XML download error:', error);
            UIModule.showStatus('Error creating XML files: ' + error.message, 'danger');
        }
    },

    async processBatchXML(csvData) {
        const zip = new JSZip();
        const results = {
            processed: 0,
            found: 0,
            notFound: []
        };

        for (const row of csvData) {
            const isbn = DataProcessor.normalizeISBN(row.ISBN || row.isbn || '');
            
            if (!isbn) continue;
            
            results.processed++;
            const book = SearchModule.findByISBN(isbn);
            
            if (book) {
                results.found++;
                try {
                    const xmlContent = this.generateXMLForBook(book);
                    zip.file(`${isbn}_c.xml`, xmlContent);
                    zip.file(`${isbn}_t.xml`, xmlContent);
                } catch (error) {
                    console.error(`Error generating XML for ISBN ${isbn}:`, error);
                    results.notFound.push(isbn);
                }
            } else {
                results.notFound.push(isbn);
            }
        }

        return { zip, results };
    }
};

// Export Module (Optimized)
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
            // Filter out items that are not available in the database
            const availableOrders = AppState.processedOrders.filter(order => order.available);
            
            if (availableOrders.length === 0) {
                UIModule.showStatus('No available items to download. All items are missing from the database.', 'warning');
                return;
            }
            
            // Re-number the filtered orders sequentially
            const renumberedOrders = availableOrders.map((order, idx) => ({
                ...order,
                lineNumber: String(idx + 1).padStart(3, '0')
            }));
            
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
        
            renumberedOrders.forEach(order => {
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
            
            const timestamp = this.getTimestamp();
            const filename = `pod_order_${timestamp}.csv`;
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        
            const omittedCount = AppState.processedOrders.length - availableOrders.length;
            let successMessage = `CSV downloaded successfully! ${availableOrders.length} items included.`;
            if (omittedCount > 0) {
                successMessage += ` ${omittedCount} unavailable item(s) omitted.`;
            }
            UIModule.showStatus(successMessage, 'success');
        } catch (error) {
            console.error('Download error:', error);
            UIModule.showStatus('Error creating CSV file: ' + error.message, 'danger');
        }
    },

    getTimestamp() {
        const now = new Date();
        return `${now.getFullYear()}_${
            String(now.getMonth() + 1).padStart(2, '0')}_${
            String(now.getDate()).padStart(2, '0')}_${
            String(now.getHours()).padStart(2, '0')}_${
            String(now.getMinutes()).padStart(2, '0')}_${
            String(now.getSeconds()).padStart(2, '0')}`;
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
                'Cover Spec': book.coverSpec || '',
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
            
            const timestamp = this.getTimestamp();
            const filename = `repository_data_${timestamp}.xlsx`;
            
            XLSX.writeFile(wb, filename);
            UIModule.showStatus(`Repository data downloaded as Excel file: ${filename}`, 'success');
            
        } catch (error) {
            console.error('Excel download error:', error);
            this.downloadRepositoryDataAsCsv();
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
                'Cover Spec': book.coverSpec || '',
                'Cover Spine': book.coverSpine || '',
                'Packing': book.packing || ''
            }));
            
            const csv = Papa.unparse(csvData);
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = window.URL.createObjectURL(blob);
            
            const a = document.createElement('a');
            a.href = url;
            const timestamp = this.getTimestamp();
            const filename = `repository_data_${timestamp}.csv`;
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

// Main Application Logic (Optimized)
async function fetchData() {
    try {
        const response = await fetch('data.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const rawData = await response.json();
        
        AppState.originalRawData = rawData;
        AppState.booksData = DataProcessor.processBookData(rawData);
        
        UIModule.showStatus(`Successfully loaded ${AppState.booksData.length} books from repository`, 'success');
        UIModule.updateRepositoryStats();
        UIModule.enableBatchXML(true);
        
    } catch (error) {
        console.error('Error fetching data:', error);
        UIModule.showStatus('Error loading book data. Please ensure data.json exists and is properly formatted.', 'danger');
        AppState.booksData = [];
        UIModule.enableBatchXML(false);
    }
}

async function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        UIModule.showStatus('Processing file...', 'info');
        
        // Parse file first to get date information
        const arrayBuffer = await file.arrayBuffer();
        const fileData = await FileProcessor.parseFileData(file, arrayBuffer);
        
        // Auto-generate order reference from Date column if available
        const orderRefField = document.getElementById('orderRef');
        if (fileData.length > 0 && !orderRefField.value) {
            const dateValue = fileData[0].Date || fileData[0].date;
            if (dateValue) {
                const generatedRef = generateOrderReference(dateValue);
                if (generatedRef) {
                    orderRefField.value = generatedRef;
                }
            }
        }
        
        const orderRef = SecurityModule.sanitizeText(orderRefField.value);
        const orderRefWarning = document.getElementById('orderRefWarning');
        
        if (!orderRef) {
            orderRefWarning.style.display = 'block';
            UIModule.showStatus('Please enter an order reference before uploading a file', 'warning');
            document.getElementById('excelFile').value = '';
            return;
        }
        
        orderRefWarning.style.display = 'none';
        
        AppState.processedOrders = FileProcessor.processOrderData(fileData, orderRef);
        AppState.sortedByPaperType = false;
        
        // Populate paper type filter dropdown
        populatePaperTypeFilter();
        
        UIModule.updatePreviewTable();
        
        const stats = AppState.processedOrders.reduce((acc, order) => {
            acc.total++;
            acc.totalQuantity += order.quantity;
            if (order.available) {
                acc.found++;
                if (order.status === 'POD Ready') acc.podReady++;
                else if (order.status === 'MPI') acc.mpi++;
            } else {
                acc.notFound++;
            }
            // Track lookup methods
            if (order.lookupMethod === 'ISBN') acc.isbnLookup++;
            else if (order.lookupMethod === 'Master Order ID') acc.masterLookup++;
            // Track consolidations
            if (order.consolidatedCount > 1) acc.consolidated++;
            return acc;
        }, { total: 0, found: 0, podReady: 0, mpi: 0, notFound: 0, isbnLookup: 0, masterLookup: 0, consolidated: 0, totalQuantity: 0 });
        
        let statusMessage = `Data loaded successfully! ${stats.found}/${stats.total} items found in repository (${stats.podReady} POD Ready, ${stats.mpi} MPI, ${stats.notFound} Not Available).`;
        
        // Add lookup method breakdown
        if (stats.isbnLookup > 0 || stats.masterLookup > 0) {
            statusMessage += ` Lookup methods: ${stats.isbnLookup} by ISBN, ${stats.masterLookup} by Master Order ID.`;
        }
        
        // Add consolidation info
        if (stats.consolidated > 0) {
            statusMessage += ` ${stats.consolidated} duplicate ISBN(s) consolidated.`;
        }
        
        UIModule.showStatus(statusMessage, 'success', true);
        UIModule.enableButtons(true);
        
    } catch (error) {
        console.error('Processing error:', error);
        UIModule.showStatus('Error processing file: ' + error.message, 'danger');
        UIModule.enableButtons(false);
        document.getElementById('excelFile').value = '';
    }
}

function generateOrderReference(dateValue) {
    if (!dateValue) return null;
    
    try {
        let date;
        
        // Handle different date formats
        if (typeof dateValue === 'string') {
            // Try to parse string date formats
            // Common formats: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, MM/DD/YYYY
            const parts = dateValue.split(/[-\/]/);
            
            if (parts.length === 3) {
                // Determine format based on first part length
                if (parts[0].length === 4) {
                    // YYYY-MM-DD format
                    date = new Date(parts[0], parseInt(parts[1]) - 1, parts[2]);
                } else if (parseInt(parts[0]) > 12) {
                    // DD-MM-YYYY or DD/MM/YYYY
                    date = new Date(parts[2], parseInt(parts[1]) - 1, parts[0]);
                } else {
                    // Assume DD-MM-YYYY or DD/MM/YYYY
                    date = new Date(parts[2], parseInt(parts[1]) - 1, parts[0]);
                }
            } else {
                date = new Date(dateValue);
            }
        } else if (typeof dateValue === 'number') {
            // Excel serial date number
            date = new Date((dateValue - 25569) * 86400 * 1000);
        } else if (dateValue instanceof Date) {
            date = dateValue;
        } else {
            return null;
        }
        
        // Validate date
        if (isNaN(date.getTime())) return null;
        
        // Format as DDMMYY
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = String(date.getFullYear()).slice(-2);
        
        return `CLAYS${day}${month}${year}`;
        
    } catch (error) {
        console.error('Error generating order reference:', error);
        return null;
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
    
    if (SecurityModule.validateISBN(rawInput)) {
        const rawFound = SearchModule.findByISBN(rawInput);
        if (rawFound) {
            foundBook = rawFound;
            searchType = 'ISBN';
        }
    }
    
    if (!foundBook) {
        const rawFound = SearchModule.findByMasterOrderId(rawInput);
        if (rawFound) {
            foundBook = rawFound;
            searchType = 'Master Order ID';
        }
    }
    
    if (foundBook) {
        const processedBook = SearchModule.createBookFromRaw(foundBook);
        const bookDataForXML = JSON.stringify(foundBook).replace(/"/g, '&quot;');
        
        UIModule.showISBNResult(`
            <div>
                <strong>Available</strong><br>
                <strong>ISBN:</strong> ${processedBook.isbn || 'Not available'}<br>
                <strong>Title:</strong> ${processedBook.title}<br>
                <strong>Master Order ID:</strong> ${processedBook.masterOrderId || 'N/A'}<br>
                <strong>Paper:</strong> ${processedBook.paperDesc}<br>
                <strong>Status:</strong> <span class="badge ${UIModule.getStatusBadgeClass(processedBook.status)}">${processedBook.status}</span><br>
                <small class="text-muted">Found by ${searchType}</small>
                <hr class="my-3">
                <button class="btn btn-primary w-100" onclick="XMLModule.downloadXMLFiles(${bookDataForXML})" title="Download XML specification files">
                    <i class="fas fa-download me-2"></i>Download XML Specification Files
                </button>
            </div>
        `, 'success');
    } else {
        const searchTypeMsg = SecurityModule.validateISBN(rawInput) ? 'ISBN' : 'Master Order ID';
        UIModule.showISBNResult(`
            <strong>Not Available</strong><br>
            <strong>${searchTypeMsg}:</strong> ${rawInput}<br>
            This title is not available in the current repository.
        `, 'danger');
    }
}

async function processBatchXML() {
    const csvUpload = document.getElementById('csvUpload');
    const file = csvUpload.files[0];
    
    if (!file) {
        UIModule.showStatus('Please select a CSV or Excel file first', 'warning');
        return;
    }

    try {
        // Validate file
        SecurityModule.validateFile(file);
        
        UIModule.showStatus('Processing file and generating XML files...', 'info');
        
        // Parse file data using the same logic as order processing
        const arrayBuffer = await file.arrayBuffer();
        const fileData = await FileProcessor.parseFileData(file, arrayBuffer);

        if (fileData.length === 0) {
            UIModule.showStatus('No data found in file', 'warning');
            return;
        }

        const { zip, results } = await XMLModule.processBatchXML(fileData);
        
        if (results.found === 0) {
            UIModule.showStatus(`No matching ISBNs found in repository out of ${results.processed} processed`, 'warning');
            return;
        }

        const timestamp = ExportModule.getTimestamp();
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        
        const url = window.URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `xml_specifications_${timestamp}.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        let statusMessage = `ZIP downloaded successfully! ${results.found}/${results.processed} ISBNs found and processed.`;
        if (results.notFound.length > 0) {
            statusMessage += ` ${results.notFound.length} ISBNs not found in repository.`;
        }
        
        UIModule.showStatus(statusMessage, 'success');
        
    } catch (error) {
        console.error('Batch XML processing error:', error);
        UIModule.showStatus('Error processing file: ' + error.message, 'danger');
    }
}

function toggleTableFilter() {
    UIModule.updatePreviewTable();
    document.getElementById('selectAll').checked = false;
    
    const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
    const showNotAvailable = document.getElementById('showNotAvailable')?.checked || false;
    const paperTypeFilter = document.getElementById('paperTypeFilter')?.value || '';
    
    if (printAsMiscellaneous && showNotAvailable) {
        if (event.target.id === 'showOnlyUnavailable') {
            document.getElementById('showNotAvailable').checked = false;
        } else {
            document.getElementById('showOnlyUnavailable').checked = false;
        }
        setTimeout(toggleTableFilter, 10);
        return;
    }
    
    const stats = AppState.processedOrders.reduce((acc, order) => {
        if (order.available && order.status === 'MPI') acc.mpi++;
        if (!order.available) acc.notAvailable++;
        return acc;
    }, { mpi: 0, notAvailable: 0 });
    
    if (printAsMiscellaneous) {
        UIModule.showStatus(`Showing ${stats.mpi} MPI items`, 'info');
    } else if (showNotAvailable) {
        UIModule.showStatus(`Showing ${stats.notAvailable} not available items`, 'info');
    } else if (paperTypeFilter) {
        const filteredCount = UIModule.getFilteredOrders().length;
        UIModule.showStatus(`Showing ${filteredCount} items with paper type: ${paperTypeFilter}`, 'info');
    } else {
        UIModule.showStatus(`Showing all ${AppState.processedOrders.length} items`, 'info');
    }
}

function populatePaperTypeFilter() {
    const paperTypeFilter = document.getElementById('paperTypeFilter');
    if (!paperTypeFilter) return;
    
    // Get unique paper types from current orders
    const paperTypes = new Set();
    AppState.processedOrders.forEach(order => {
        if (order.paperDesc && order.paperDesc !== 'Not specified') {
            paperTypes.add(order.paperDesc);
        }
    });
    
    // Sort alphabetically
    const sortedPaperTypes = Array.from(paperTypes).sort();
    
    // Clear existing options except the first "All Paper Types"
    paperTypeFilter.innerHTML = '<option value="">All Paper Types</option>';
    
    // Add paper type options
    sortedPaperTypes.forEach(paperType => {
        const option = document.createElement('option');
        option.value = paperType;
        option.textContent = paperType;
        paperTypeFilter.appendChild(option);
    });
}

function toggleAllCheckboxes() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const selectAllCheckbox = document.getElementById('selectAll');
    checkboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
    });
}

// Global function exports for HTML onclick handlers
function clearAll() {
    OrderProcessor.clearAll();
}

function downloadCsv() {
    ExportModule.downloadCsv();
}

function downloadRepositoryData() {
    ExportModule.downloadRepositoryData();
}

function downloadTemplate() {
    ExportModule.downloadTemplate();
}

function deleteSelected() {
    OrderProcessor.deleteSelected();
}

function sortByPaperType() {
    if (AppState.processedOrders.length === 0) {
        UIModule.showStatus('No data to sort', 'warning');
        return;
    }
    
    // Toggle sort state
    AppState.sortedByPaperType = !AppState.sortedByPaperType;
    
    if (AppState.sortedByPaperType) {
        // Sort by paper type, then by description
        AppState.processedOrders.sort((a, b) => {
            const paperA = (a.paperDesc || 'Not specified').toLowerCase();
            const paperB = (b.paperDesc || 'Not specified').toLowerCase();
            
            if (paperA === paperB) {
                return a.description.localeCompare(b.description);
            }
            return paperA.localeCompare(paperB);
        });
        
        UIModule.showStatus('Orders sorted by paper type', 'info');
    } else {
        // Sort by line number (restore original order)
        AppState.processedOrders.sort((a, b) => {
            return parseInt(a.lineNumber) - parseInt(b.lineNumber);
        });
        
        UIModule.showStatus('Orders restored to original order', 'info');
    }
    
    UIModule.updatePreviewTable();
}

// Make functions globally available
window.clearAll = clearAll;
window.downloadCsv = downloadCsv;
window.downloadRepositoryData = downloadRepositoryData;
window.downloadTemplate = downloadTemplate;
window.searchISBN = searchISBN;
window.deleteSelected = deleteSelected;
window.toggleTableFilter = toggleTableFilter;
window.toggleAllCheckboxes = toggleAllCheckboxes;
window.processBatchXML = processBatchXML;
window.sortByPaperType = sortByPaperType;

// Initialize the application
function initializeApp() {
    const excelFile = document.getElementById('excelFile');
    const isbnSearch = document.getElementById('isbnSearch');
    const paperTypeFilter = document.getElementById('paperTypeFilter');
    
    if (excelFile) {
        excelFile.addEventListener('change', handleFileSelect);
    }
    
    if (isbnSearch) {
        isbnSearch.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                searchISBN();
            }
        });
        
        isbnSearch.addEventListener('input', function(e) {
            const value = e.target.value;
            
            if (/[a-zA-Z]/.test(value) && !SecurityModule.validateISBN(value)) {
                e.target.value = value.toUpperCase();
            }
        });
    }
    
    // Add explicit event listener for paper type filter (Chrome compatibility)
    if (paperTypeFilter) {
        paperTypeFilter.addEventListener('change', function(e) {
            toggleTableFilter();
        });
    }
    
    fetchData();
}

// Wait for DOM to be fully loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}