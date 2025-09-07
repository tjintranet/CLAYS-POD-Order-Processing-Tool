# CLAYS POD Order Processing Tool

A streamlined web-based application for processing print-on-demand orders for Clays Ltd. Features complete order processing, real-time repository validation, ISBN/Master Order ID lookup with XML specification downloads, batch XML generation, and an integrated repository statistics dashboard with download functionality.

## 🚀 Features

### **Order Processing**
- **Excel/CSV Upload**: Process `.xlsx`, `.xls`, `.csv` order files up to 10MB
- **Excel Template**: Download pre-formatted Excel template for consistent order format
- **Real-time Validation**: Check ISBNs against live repository database with status information
- **Smart Filtering**: 
  - **POD Ready Only**: Show only items available for print-on-demand
  - **Miscellaneous Print Items (MPI)**: Show only MPI status items
  - **Show Not Available Items**: Display items not found in repository
- **Bulk Operations**: Select and delete multiple rows with checkbox selection
- **CSV Export**: Generate Clays POD-compliant CSV files with correct format

### **ISBN/Master Order ID Lookup with XML Generation**
- **Dual Search**: Search by ISBN or Master Order ID with Enter key support
- **Comprehensive Results**: Shows ISBN, title, Master Order ID, status, and paper description
- **XML Download**: Download complete XML specification files (`{ISBN}_c.xml` and `{ISBN}_t.xml`) for found items
- **Batch XML Processing**: Upload CSV/Excel file with ISBN column to generate ZIP archive of all XML files
- **Status Indicators**: 
  - 🟢 **POD Ready**: Available for print-on-demand
  - 🟡 **MPI**: Miscellaneous Print Item
  - 🔴 **Not Available**: Item not in repository catalog
- **Format Flexibility**: Accepts ISBNs with or without dashes/spaces
- **Auto-hide Results**: Success results disappear after 10 seconds

### **XML Specification Generation**
- **Automatic Generation**: Creates XML files matching Clays printing specifications
- **Dual File Output**: Downloads both `{ISBN}_c.xml` and `{ISBN}_t.xml` files simultaneously
- **Complete Specifications**: Includes book dimensions, materials, binding, and extent information
- **Secure Processing**: XML content is properly sanitized and validated
- **Professional Format**: UTF-8 encoded XML with proper structure for production use
- **Batch Processing**: Generate ZIP archives containing XML files for multiple ISBNs

### **Repository Statistics Dashboard**
- **Real-time Overview**: Total titles, POD Ready count, MPI count, and duplicate detection
- **Repository Download**: Export complete repository data as Excel file with all fields
- **Visual Design**: Clean, professional statistics display integrated on main page
- **Duplicate Detection**: Accurately identifies and counts duplicate ISBN entries

### **Security & Data Handling** 🛡️
- **Input Sanitization**: Comprehensive validation of all user inputs including XML-specific sanitization
- **File Validation**: Size limits (10MB), type checking, and format verification
- **Content Security Policy**: XSS protection and resource validation
- **Error Handling**: Safe error messages with graceful degradation

## 📋 Requirements

### **Browser Compatibility**
- Modern web browsers (Chrome 80+, Firefox 75+, Safari 13+, Edge 80+)
- JavaScript enabled
- File API and Clipboard API support

### **Required Files**
```
project-root/
├── index.html              # Main application interface
├── script.js               # Application logic
├── data.json               # Repository database (required)
└── order_template.xlsx     # Excel template for orders
```

### **Repository Database Format**
The `data.json` file must contain an array of book objects with the complete structure for XML generation:
```json
[
  {
    "ISBN": 9780140175936,
    "Master Order ID": "SA1657",
    "TITLE": "CLEOPATRA'S SISTER   (11)",
    "Trim Height": 198,
    "Trim Width": 129,
    "Bind Style": "Limp",
    "Extent": 288,
    "Paper Desc": "Holmen Bulky 52 gsm",
    "Cover Spec Code 1": "C400P2",
    "Cover Spine": 18,
    "Packing": "Pack (64) (8) Base.",
    "Status": "POD Ready"
  }
]
```

**Required Fields for XML Generation:**
- `ISBN` or `isbn`: Book identifier
- `TITLE`, `Title`, or `title`: Book title
- `Trim Height`: Book height in mm
- `Trim Width`: Book width in mm
- `Cover Spine` or `Spine Size`: Spine thickness in mm
- `Paper Desc` or `Paper Description`: Paper specification
- `Bind Style` or `Binding Style`: Binding method
- `Extent` or `Page Extent`: Number of pages

## 🛠️ Installation

### **Basic Setup**
1. **Download** all project files to a local directory
2. **Create** `data.json` with your book repository (see format above)
3. **Add** `order_template.xlsx` Excel template file
4. **Open** `index.html` in a web browser

### **GitHub Pages Deployment**
1. **Create** a new GitHub repository
2. **Upload** all files to the repository
3. **Enable** GitHub Pages in repository settings
4. **Access** via the provided GitHub Pages URL

### **Local Development Server**
For local file restrictions, use a simple HTTP server:
```bash
# Python 3
python -m http.server 8000

# Node.js (with http-server)
npx http-server

# PHP
php -S localhost:8000
```

## 📖 Usage Guide

### **Getting Started**

#### **Initial Setup**
1. Open the application in your web browser
2. The tool automatically loads repository from `data.json`
3. View repository statistics in the dashboard at the top

### **Order Processing Workflow**

#### **1. Using the Excel Template**
1. **Download Template**: Click "Download Excel Template" under the file upload
2. **Fill Template**: Add ISBN and Quantity columns for your orders
3. **Save File**: Save as Excel (.xlsx) or CSV format

#### **2. Processing Orders**
1. **Enter Order Reference** (required field)
2. **Upload File** using the template or CSV with ISBN and Qty columns
3. **Review Results** in the preview table with status indicators
4. **Filter Items** using available filter options:
   - **POD Ready Only**: Shows only POD Ready status items
   - **Miscellaneous Print Items (MPI)**: Shows only MPI status items
   - **Show Not Available Items**: Shows items not in repository catalog
5. **Export CSV** for Clays POD processing

#### **3. Supported File Formats**
Your order files should contain:
| Column | Description | Example |
|--------|-------------|---------|
| ISBN | Book identifier (10-13 digits) | 9781234567890 |
| Qty | Quantity to order | 5 |

**Supported variations:**
- Column names: `ISBN`, `isbn`, `Qty`, `qty`, `Quantity`, `quantity`
- ISBN formats: With/without dashes, scientific notation handling
- File types: `.xlsx`, `.xls`, `.csv` (max 10MB)

#### **4. Status Understanding**
- 🟢 **POD Ready**: Items available for print-on-demand production
- 🟡 **MPI**: Miscellaneous Print Items requiring special handling
- 🔴 **Not Available**: Items not found in current repository catalog

#### **5. Managing Orders**
- **Individual Deletion**: Click trash icon in Action column
- **Bulk Selection**: Use checkboxes to select multiple rows
- **Delete Selected**: Remove multiple items at once
- **Smart Filtering**: Focus on specific item types or statuses

### **ISBN/Master Order ID Lookup with XML Downloads**

#### **Individual Search**
1. Enter ISBN (e.g., `9780140175936`) or Master Order ID (e.g., `SA1657`)
2. Click "Search Repository" or press Enter
3. View comprehensive results including status badge and paper description
4. Click "Download XML Specification Files" to get both XML files

#### **Batch XML Generation**
1. **Upload File**: Select CSV or Excel file with ISBN column
2. **Process**: Click "Download XML ZIP" 
3. **Download**: Receive ZIP archive containing all XML specification files
4. **Results**: View processing summary with found/not found counts

#### **Search Results with XML Download**
- **Available Items**: Shows complete book information with status badge
- **XML Download**: Full-width "Download XML Specification Files" button
- **Dual File Download**: Automatically downloads both `{ISBN}_c.xml` and `{ISBN}_t.xml`
- **Not Available Items**: Clear indication that item is not in repository
- **Auto-hide**: Success results disappear after 10 seconds

#### **XML File Structure**
Downloaded XML files contain:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<book>
    <basic_info>
        <isbn>9780002711845</isbn>
        <title>LIVING (00)</title>
    </basic_info>
    <specifications>
        <dimensions>
            <trim_height>216</trim_height>
            <trim_width>135</trim_width>
            <spine_size>17</spine_size>
        </dimensions>
        <materials>
            <paper_type>Holmen Cream 65 gsm</paper_type>
            <binding_style>Limp</binding_style>
            <lamination>No Lamination</lamination>
        </materials>
        <page_extent>240</page_extent>
    </specifications>
</book>
```

### **Repository Management**

#### **Statistics Dashboard**
The dashboard provides four key metrics:
- **Total Titles**: Complete repository count
- **POD Ready**: Items available for print-on-demand
- **MPI Items**: Items requiring miscellaneous print handling
- **Duplicates**: Number of duplicate ISBN entries detected

#### **Repository Export**
- **Download Button**: Located in repository statistics header
- **Excel Format**: Primary export as .xlsx with auto-sized columns
- **CSV Fallback**: Automatic fallback if Excel export fails
- **Complete Data**: Includes all repository fields and metadata
- **Timestamped Files**: Automatic filename with date/time

## 🔧 Configuration

### **Customer Information**
Update Clays POD details in `script.js`:
```javascript
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
    "phone": "01986 893 211"
};
```

## 📤 CSV Output Format

Generated CSV files follow Clays POD specifications:

### **Header Line (HDR)**
```csv
HDR,ORDER123,20240605,Clays POD,Clays Ltd,Popson Street,,Bungay,Suffolk,UK,NR35 1ED,GB,01986 893 211,
```

### **Detail Lines (DTL)**
```csv
DTL,ORDER123,001,9781234567890,5
DTL,ORDER123,002,9780987654321,3
```

**Structure:**
- **HDR Line**: Header with order details and Clays Ltd address information
- **DTL Lines**: Detail lines with exactly 5 fields:
  1. DTL marker
  2. Order reference
  3. Line number (001, 002, etc.)
  4. ISBN (13 digits, zero-padded)
  5. Quantity

## 🛡️ Security Features

### **Content Security Policy**
- Prevents XSS attacks through resource control
- Blocks unauthorized script execution
- Validates all external resource loading

### **Input Validation**
- **File Type Checking**: Only allows specified file extensions (.xlsx, .xls, .csv)
- **Size Limits**: 10MB maximum for all file uploads
- **ISBN Validation**: Ensures proper format before processing
- **Text Sanitization**: Removes potentially dangerous characters including XML-specific sanitization
- **Quantity Validation**: Reasonable limits and numeric validation

### **Error Handling**
- **Safe Error Messages**: No sensitive information disclosure
- **Graceful Degradation**: Continues operation despite individual component failures
- **User Feedback**: Clear status messages for all operations

## 🔍 Troubleshooting

### **Common Issues**

#### **Application Won't Load**
- **Browser Console**: Press F12, check Console tab for errors
- **File Location**: Ensure all files are in the same directory
- **JavaScript Enabled**: Verify browser settings allow JavaScript
- **HTTPS**: Use HTTPS for production deployment

#### **"Error loading book data"**
- **File Exists**: Confirm `data.json` is present in same directory
- **Valid JSON**: Check syntax with JSON validator
- **Proper Format**: Ensure array structure with complete book objects
- **Required Fields**: ISBN, Master Order ID, TITLE, Status fields required for basic functionality
- **XML Fields**: Additional fields required for XML generation (see Repository Database Format)
- **File Permissions**: Ensure web server can read the file

#### **File Upload Problems**
- **File Size**: Maximum 10MB limit for all uploads
- **File Type**: Only `.xlsx`, `.xls`, `.csv` files accepted
- **Column Names**: Use `ISBN`/`isbn` and `Qty`/`quantity` columns
- **Order Reference**: Must be entered before file upload
- **Template Format**: Use provided Excel template for best results

#### **Template Download Issues**
- **File Missing**: Ensure `order_template.xlsx` exists in same directory as `index.html`
- **Server Permissions**: Check that web server can serve .xlsx files
- **Pop-up Blockers**: Disable pop-up blockers for the application domain
- **Browser Settings**: Ensure downloads are allowed

#### **XML Download Issues**
- **Missing Data**: Ensure repository contains required fields for XML generation
- **File Generation**: Check browser console for XML generation errors
- **Download Blocked**: Verify browser allows multiple file downloads
- **Pop-up Blockers**: Disable pop-up blockers for the application domain
- **Field Mapping**: Ensure JSON uses expected field names (TITLE, Trim Height, etc.)

#### **Batch XML Processing Issues**
- **File Format**: Ensure CSV/Excel file contains ISBN column
- **Large Files**: Process files with many ISBNs may take time
- **ZIP Generation**: Browser must support JSZip for batch processing
- **Memory Limits**: Very large batches may hit browser memory limits

#### **Books Showing "Not Available"**
- **ISBN Matching**: Ensure exact match with repository ISBNs
- **Format Consistency**: Check for leading zeros or formatting differences
- **Database Update**: Verify `data.json` contains expected entries
- **Case Sensitivity**: Master Order IDs are case-insensitive
- **Data Types**: Ensure ISBNs are stored as numbers or strings, not other formats

#### **Status Display Issues**
- **Missing Status**: Items without Status field default to "POD Ready"
- **Valid Values**: Status should be "POD Ready", "MPI", or custom status
- **Badge Colors**: Green=POD Ready, Yellow=MPI, Red=Not Available

#### **CSV Export Problems**
- **Format Issues**: DTL lines should have exactly 5 comma-separated fields
- **Character Encoding**: Files saved as UTF-8
- **Line Endings**: Uses standard CSV line endings
- **Special Characters**: Text is properly escaped in CSV output

#### **Search Not Working**
- **ISBN Format**: Try with or without dashes/spaces
- **Master Order ID**: Ensure exact match (case-insensitive)
- **Data Loading**: Ensure repository loaded successfully
- **Field Names**: Check that JSON uses expected field names

#### **Duplicate Count Issues**
- **Empty ISBNs**: Empty ISBN fields are excluded from duplicate detection
- **Data Processing**: ISBNs are normalized before comparison
- **Field Names**: Ensure ISBN field is correctly named in JSON

### **Advanced Troubleshooting**

#### **Performance Issues**
- **Large Files**: Break large uploads into smaller batches (< 5MB recommended)
- **Browser Memory**: Close other tabs if processing large repositories
- **Network**: Ensure stable internet for CDN resource loading
- **Repository Size**: Tested with 10,000+ entries

#### **Browser Compatibility**
- **Modern Browsers**: Use recent versions of Chrome, Firefox, Safari, Edge
- **Feature Support**: Verify File API and Clipboard API availability
- **JavaScript Version**: Ensure ES6+ support
- **Console Errors**: Check browser console for specific error messages

## 📄 Maintenance

### **Regular Tasks**
- **Repository Updates**: Keep `data.json` current with new titles and status changes
- **XML Field Validation**: Ensure new entries contain all required fields for XML generation
- **Template Updates**: Update Excel template if column requirements change
- **Backup Data**: Regular backups of repository database
- **Library Updates**: Monitor CDN dependencies for security updates

### **Monitoring**
- **Error Logs**: Check browser console for recurring issues
- **Usage Patterns**: Monitor which features are most used (order processing vs XML downloads)
- **Performance**: Watch for slow uploads or processing times
- **Status Accuracy**: Verify status indicators match actual production capabilities
- **XML Generation**: Test XML downloads with sample data regularly

### **Updates**
- **Dependencies**: Keep external libraries updated for security
- **Feature Requests**: Document user feedback for improvements
- **Version Control**: Use Git for tracking changes
- **Documentation**: Keep README.md updated with any changes

## 📊 Technical Specifications

### **Dependencies**
- **Bootstrap 5.3.2**: UI framework and responsive design
- **Font Awesome 6.4.0**: Icon library
- **SheetJS 0.18.5**: Excel file processing
- **PapaParse 5.4.1**: CSV parsing and generation
- **JSZip 3.10.1**: ZIP file generation for batch XML downloads

### **Browser APIs Used**
- **File API**: For file upload and processing
- **Blob API**: For file downloads including XML generation
- **Fetch API**: For loading repository data
- **URL API**: For file download handling

### **Performance Characteristics**
- **File Processing**: Up to 10MB files supported
- **Repository Size**: Tested with 10,000+ book entries
- **Response Time**: Near-instantaneous for local operations
- **Memory Usage**: Optimized for large datasets
- **XML Generation**: Real-time generation and download
- **Batch Processing**: Handles hundreds of XML files efficiently

### **Data Structure Requirements**
- **Complete JSON**: Repository must include full data.json structure
- **Required Fields**: ISBN, Master Order ID, TITLE, Status minimum
- **XML Fields**: Additional fields required for XML generation (dimensions, materials, etc.)
- **Status Values**: "POD Ready", "MPI", or custom status strings
- **ISBN Format**: 10-13 digit numbers, automatically normalized
- **Flexible Field Names**: Supports various ISBN and title field name variations

## 🤝 Support

### **Getting Help**
- **Documentation**: Review this README for detailed guidance
- **Browser Console**: Check for error messages and warnings (F12)
- **Test Data**: Use sample data to isolate issues
- **Version Check**: Ensure all files are from the same version

### **Reporting Issues**
When reporting problems, include:
- Browser type and version
- Error messages from console (F12 → Console tab)
- File types and sizes being processed
- Steps to reproduce the issue
- Sample data (if applicable)
- Expected vs actual behavior
- Whether issue is with order processing, XML downloads, or batch processing

### **Feature Requests**
The application is designed for extensibility. Common enhancement areas:
- Additional status types and workflows
- Enhanced reporting capabilities
- Integration with external systems
- Batch processing optimizations
- Custom CSV export formats
- Advanced XML customization options

## 📝 Version History

### **Current Version** - Enhanced Batch XML Processing
- **Enhanced**: Batch XML generation with ZIP download functionality
- **Added**: CSV/Excel upload for batch XML processing
- **Improved**: XML button styling - now full-width primary button at bottom of search results
- **Added**: Proper XML sanitization for security
- **Enhanced**: Better field name detection for XML generation (TITLE variations)
- **Improved**: Error handling for XML generation and downloads
- **Enhanced**: User experience with cleaner search result layout
- **Added**: JSZip dependency for batch processing
- **Fixed**: Template download functionality
- **Enhanced**: Repository export with complete data structure
- **Improved**: Duplicate detection accuracy

### **Previous Versions**
- **v5.0**: Streamlined interface focused on core order processing workflow
- **v4.0**: Status field support with POD Ready/MPI/Not Available indicators
- **v3.0**: Complete security enhancement with comprehensive input validation
- **v2.0**: Enhanced security & ISBN search functionality
- **v1.0**: Initial release with core Excel upload, validation, and CSV export

## 📄 License

This application is designed specifically for Clays Ltd print-on-demand operations. All configuration and branding elements are customized for Clays POD workflow requirements with status-aware processing capabilities and XML specification generation.

## 🎯 Quick Start Guide

1. **Setup**: Place `index.html`, `script.js`, `data.json`, and `order_template.xlsx` in same directory
2. **Open**: Launch `index.html` in web browser
3. **Template**: Download Excel template using the button
4. **Prepare**: Fill template with ISBN and Quantity columns
5. **Process**: Enter order reference, upload file, review results
6. **Export**: Download CSV for Clays POD processing
7. **XML Downloads**: Search for individual ISBNs or upload batch file for XML specification downloads

---

**For technical support or customization requests, please refer to the troubleshooting section above or consult your system administrator.**