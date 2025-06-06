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

async function fetchData() {
    try {
        console.log('Attempting to fetch data.json...');
        const response = await fetch('data.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const rawData = await response.json();
        console.log('Raw data loaded:', rawData);
        
        // Process and normalize the data
        booksData = rawData.map(item => {
            // Convert numeric codes to strings and pad to 13 digits
            const code = String(item.code).padStart(13, '0');
            
            return {
                code: code,
                description: item.description || 'No description available',
                setupdate: item.setupdate || '' // Default to empty string if not provided
            };
        });
        
        console.log('Processed books data:', booksData);
        console.log(`Successfully loaded ${booksData.length} books`);
        showStatus(`Successfully loaded ${booksData.length} books from inventory`, 'success');
        
    } catch (error) {
        console.error('Error fetching data:', error);
        showStatus('Error loading book data: ' + error.message + '. Please ensure data.json exists and is properly formatted.', 'danger');
        
        // Set empty array to prevent further errors
        booksData = [];
    }
}

document.getElementById('excelFile').addEventListener('change', handleFileSelect);

async function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    const orderRef = document.getElementById('orderRef').value.trim();
    const orderRefWarning = document.getElementById('orderRefWarning');
    
    if (!orderRef) {
        orderRefWarning.style.display = 'block';
        showStatus('Please enter an order reference before uploading a file', 'warning');
        document.getElementById('excelFile').value = '';
        return;
    }
    
    orderRefWarning.style.display = 'none';

    try {
        showStatus('Processing file...', 'info');
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        console.log('Excel data loaded:', excelData);

        // Create a map for faster lookup - using normalized ISBN as key
        const booksMap = new Map();
        booksData.forEach(item => {
            booksMap.set(item.code, item);
        });

        console.log('Books map created:', booksMap);

        processedOrders = excelData.map((row, index) => {
            console.log('Processing Excel row:', row);
            
            // Handle ISBN - convert to string and normalize
            let isbn = String(row.ISBN || row.isbn || '').trim();
            
            // Handle scientific notation in Excel
            if (isbn.includes('e') || isbn.includes('E')) {
                isbn = Number(isbn).toFixed(0);
            }
            
            // Remove all non-digits and pad to 13 digits
            isbn = isbn.replace(/\D/g, '').padStart(13, '0');
            
            // Handle quantity
            const quantity = parseInt(row.Qty || row.qty || row.Quantity || row.quantity || 0);
            
            console.log('Normalized ISBN:', isbn, 'Quantity:', quantity);
            
            // Look up book in inventory
            const stockItem = booksMap.get(isbn);
            console.log('Stock item found:', stockItem);
            
            return {
                lineNumber: String(index + 1).padStart(3, '0'),
                orderRef,
                isbn,
                description: stockItem?.description || 'Not Found',
                quantity,
                available: !!stockItem,
                setupDate: stockItem?.setupdate || ''
            };
        });

        console.log('Processed orders:', processedOrders);
        updatePreviewTable();
        
        const availableCount = processedOrders.filter(order => order.available).length;
        const totalCount = processedOrders.length;
        showStatus(`Data loaded successfully! ${availableCount}/${totalCount} items found in inventory.`, 'success');
        enableButtons(true);
        
    } catch (error) {
        console.error('Processing error:', error);
        showStatus('Error processing file. Please check the format and try again. Error: ' + error.message, 'danger');
        enableButtons(false);
    }
}

function updatePreviewTable() {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';

    if (processedOrders.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center">No data loaded</td></tr>';
        return;
    }

    const showOnlyUnavailable = document.getElementById('showOnlyUnavailable').checked;
    let filteredOrders = processedOrders;

    if (showOnlyUnavailable) {
        filteredOrders = processedOrders.filter(order => !order.available);
    }

    if (filteredOrders.length === 0 && showOnlyUnavailable) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-success">All items are available in inventory!</td></tr>';
        return;
    }

    filteredOrders.forEach((order, index) => {
        // Use the original index from processedOrders for operations
        const originalIndex = processedOrders.indexOf(order);
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <input type="checkbox" class="row-checkbox" data-index="${originalIndex}">
            </td>
            <td>
                <button class="btn btn-danger btn-sm" onclick="deleteRow(${originalIndex})">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
            <td>${order.lineNumber}</td>
            <td>${order.isbn}</td>
            <td>${order.description}</td>
            <td>${order.quantity}</td>
            <td>
                <span class="badge ${order.available ? 'bg-success' : 'bg-danger'}">
                    ${order.available ? 'Available' : 'Not Found'}
                </span>
            </td>
        `;
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

function downloadCsv() {
    if (processedOrders.length === 0) {
        showStatus('No data to download', 'warning');
        return;
    }

    try {
        const formattedDate = formatDate();
        const orderRef = document.getElementById('orderRef').value;
    
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
    
        // Safely append an extra blank column for CSV without affecting other functions
        const finalCsvRow = [...csvRow, ''];
    
        const csvContent = [finalCsvRow];
    
        processedOrders.forEach(order => {
            const dtlRow = Array(customerConfig.csvStructure.length).fill('');
            dtlRow[0] = 'DTL';
            dtlRow[1] = orderRef; // Add order reference to DTL rows
            dtlRow[2] = order.lineNumber;
            dtlRow[3] = order.isbn;
            dtlRow[4] = order.quantity.toString();
            // Add empty column to match header
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
        tableHtml += `
            <tr>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${order.lineNumber}</td>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${order.isbn}</td>
                <td style="padding: 8px; border: 1px solid #dee2e6;">${order.description}</td>
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

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.className = `alert alert-${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000); // Increased timeout for better visibility
    }
}

function toggleTableFilter() {
    updatePreviewTable();
    
    // Update the "Select All" checkbox to unchecked when filter changes
    document.getElementById('selectAll').checked = false;
    
    const showOnlyUnavailable = document.getElementById('showOnlyUnavailable').checked;
    const unavailableCount = processedOrders.filter(order => !order.available).length;
    
    if (showOnlyUnavailable) {
        showStatus(`Showing ${unavailableCount} unavailable items`, 'info');
    } else {
        showStatus(`Showing all ${processedOrders.length} items`, 'info');
    }
}

// Initialize the application
console.log('Initializing application...');
fetchData();