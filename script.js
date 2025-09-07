// Application State Management
const AppState = {
    booksData: [],
    processedOrders: [],
    originalRawData: null,
    booksMap: new Map(),
    masterOrderMap: new Map()
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
    showStatus(message, type) {
        const statusDiv = document.getElementById('status');
        statusDiv.className = `alert alert-${type}`;
        statusDiv.textContent = message;
        statusDiv.style.display = 'block';
        
        if (type === 'success' || type === 'info') {
            setTimeout(() => {
                statusDiv.style.display = 'none';
            }, 15000);
        }
    },

    enableButtons(enabled) {
        const buttonIds = ['clearBtn', 'downloadBtn', 'deleteSelectedBtn'];
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
            tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted"><i class="fas fa-inbox"></i> No data loaded</td></tr>';
            return;
        }

        const filteredOrders = this.getFilteredOrders();

        if (filteredOrders.length === 0) {
            const message = this.getFilterMessage();
            tbody.innerHTML = `<tr><td colspan="7" class="text-center text-success">${message}</td></tr>`;
            return;
        }

        const fragment = document.createDocumentFragment();
        filteredOrders.forEach(order => {
            const originalIndex = AppState.processedOrders.indexOf(order);
            const tr = this.createTableRow(order, originalIndex);
            fragment.appendChild(tr);
        });
        
        tbody.appendChild(fragment);
    },

    getFilteredOrders() {
        const showOnlyPodReady = document.getElementById('showOnlyPodReady')?.checked || false;
        const printAsMiscellaneous = document.getElementById('showOnlyUnavailable').checked;
        const showNotAvailable = document.getElementById('showNotAvailable')?.checked || false;
        
        if (showOnlyPodReady) {
            return AppState.processedOrders.filter(order => order.available && order.status === 'POD Ready');
        } else if (printAsMiscellaneous) {
            return AppState.processedOrders.filter(order => order.available && order.status === 'MPI');
        } else if (showNotAvailable) {
            return AppState.processedOrders.filter(order => !order.available);
        }
        
        return AppState.processedOrders;
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
        
        // Build row HTML efficiently
        tr.innerHTML = `
            <td><input type="checkbox" class="row-checkbox" data-index="${originalIndex}"></td>
            <td><button class="btn btn-danger btn-sm" onclick="OrderProcessor.deleteRow(${originalIndex})"><i class="fas fa-trash"></i></button></td>
            <td>${order.lineNumber}</td>
            <td>${order.isbn}</td>
            <td>${order.description}</td>
            <td>${order.quantity}</td>
            <td><span class="badge ${this.getStatusBadgeClass(order.status)}">${order.available ? order.status : 'Not Available'}</span></td>
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
        return excelData.map((row, index) => {
            let isbn = String(row.ISBN || row.isbn || '').trim();
            
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            if (!SecurityModule.validateISBN(isbn)) {
                isbn = '';
            } else {
                isbn = DataProcessor.normalizeISBN(isbn);
            }
            
            const rawQuantity = row.Qty || row.qty || row.Quantity || row.quantity || 0;
            const quantity = SecurityModule.validateQuantity(rawQuantity) ? parseInt(rawQuantity) : 0;
            
            const stockItem = AppState.booksMap.get(isbn);
            
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

// Order Processing Module (Optimized)
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

// Export Module (Optimized and Copy-to-Clipboard removed)
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
            
            const timestamp = this.getTimestamp();
            const filename = `pod_order_${timestamp}.csv`;
            
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
        
        const stats = AppState.processedOrders.reduce((acc, order) => {
            acc.total++;
            if (order.available) {
                acc.found++;
                if (order.status === 'POD Ready') acc.podReady++;
                else if (order.status === 'MPI') acc.mpi++;
            } else {
                acc.notFound++;
            }
            return acc;
        }, { total: 0, found: 0, podReady: 0, mpi: 0, notFound: 0 });
        
        UIModule.showStatus(`Data loaded successfully! ${stats.found}/${stats.total} items found in repository (${stats.podReady} POD Ready, ${stats.mpi} MPI, ${stats.notFound} Not Available).`, 'success');
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
        
        isbnSearch.addEventListener('input', function(e) {
            const value = e.target.value;
            
            if (/[a-zA-Z]/.test(value) && !SecurityModule.validateISBN(value)) {
                e.target.value = value.toUpperCase();
            }
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