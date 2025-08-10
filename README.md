# CLAYS POD Order Processing & Repo Management Tool

A comprehensive web-based application for processing print-on-demand orders and managing inventory for Clays Ltd. Features complete order processing, real-time repo validation, ISBN/Master Order ID lookup, and secure repo management with password protection.

## 🚀 Features

### **Order Processing**
- **Excel/CSV Upload**: Process `.xlsx`, `.xls`, `.csv` order files up to 10MB
- **Real-time Validation**: Check ISBNs against live repo database with status information
- **Smart Filtering**: 
  - **Print as Miscellaneous Items (MPI)**: Show only MPI status items
  - **Show Not Available Items**: Display items not found in repo
- **Bulk Operations**: Select and delete multiple rows, copy to clipboard
- **CSV Export**: Generate Clays POD-compliant CSV files
- **Template Download**: Get pre-formatted Excel template for orders

### **ISBN/Master Order ID Management**
- **Dual Search**: Search by ISBN or Master Order ID
- **Comprehensive Results**: Shows ISBN, title, Master Order ID, and status
- **Status Indicators**: 
  - 🟢 **POD Ready**: Available for print-on-demand
  - 🟡 **MPI**: Miscellaneous Print Item
  - 🔴 **Not Available**: Item not in repo catalog
- **Format Flexibility**: Accepts ISBNs with or without dashes/spaces
- **Keyboard Shortcuts**: Press Enter to trigger searches

### **Complete Repo Management** 🔒
- **JSON-Only Upload**: Upload complete data.json format files
- **Add New Titles**: Expand repo with full data structure preservation
- **Remove Titles**: Individual or bulk removal of existing titles
- **Dual Preview**: Separate tables for additions and removals
- **Statistics Dashboard**: 6-metric real-time repo overview
- **Change Management**: Preview all modifications before applying
- **Export Options**: Download updated repo + backup current database

### **Security & Access Control** 🛡️
- **Password Protection**: Secure masked password input for repo management
- **Session Management**: 1-hour sessions with automatic extension
- **Visual Session Timer**: Real-time countdown with expiration warnings
- **Secure Logout**: One-click session termination
- **Content Security Policy**: XSS protection and resource validation
- **Input Sanitization**: Comprehensive validation of all user inputs

## 📋 Requirements

### **Browser Compatibility**
- Modern web browsers (Chrome 80+, Firefox 75+, Safari 13+, Edge 80+)
- JavaScript enabled
- File API and Clipboard API support

### **Required Files**
```
project-root/
├── index.html              # Main application interface
├── script.js               # Enhanced application logic
├── data.json               # Repo database (required)
├── order_template.xlsx     # Excel template for orders
├── favicon-32x32.png       # Favicon (optional)
└── apple-touch-icon.png    # Apple touch icon (optional)
```

### **Repo Database Format**
The `data.json` file must contain an array of book objects with the complete structure:
```json
[
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
]
```

## 🛠️ Installation

### **Basic Setup**
1. **Download** all project files to a local directory
2. **Create** `data.json` with your book repo (see format above)
3. **Place** `order_template.xlsx` in the same directory (optional)
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
2. The tool automatically loads repo from `data.json`
3. Download the order template for standardized Excel format

#### **Default Credentials**
- **Repo Management Password**: `admin123`
- **Session Duration**: 1 hour with automatic extension
- **Access Level**: All repo modification functions

### **Order Processing Workflow**

#### **1. Basic Order Processing**
1. **Enter Order Reference** (required field)
2. **Upload File** with ISBN and Qty columns (Excel/CSV supported)
3. **Review Results** in the preview table with status indicators
4. **Filter Items** using available filter options:
   - **Print as Miscellaneous Items (MPI)**: Shows only MPI status items
   - **Show Not Available Items**: Shows items not in repo catalog
5. **Export CSV** for Clays POD processing

#### **2. Supported File Formats**
Your order files should contain:
| Column | Description | Example |
|--------|-------------|---------|
| ISBN | Book identifier (10-13 digits) | 9781234567890 |
| Qty | Quantity to order | 5 |

**Supported variations:**
- Column names: `ISBN`, `isbn`, `Qty`, `qty`, `Quantity`, `quantity`
- ISBN formats: With/without dashes, scientific notation
- File types: `.xlsx`, `.xls`, `.csv` (max 10MB)

#### **3. Status Understanding**
- 🟢 **POD Ready**: Items available for print-on-demand production
- 🟡 **MPI**: Miscellaneous Print Items requiring special handling
- 🔴 **Not Available**: Items not found in current repo catalog

#### **4. Managing Orders**
- **Individual Deletion**: Click trash icon in Action column
- **Bulk Selection**: Use checkboxes to select multiple rows
- **Copy to Clipboard**: Export table as HTML for external use
- **Smart Filtering**: Focus on specific item types or statuses

### **ISBN/Master Order ID Lookup**

#### **Dual Search Capability**
1. Enter ISBN (e.g., `9780140175936`) or Master Order ID (e.g., `SA1657`)
2. Click "Search Repo" or press Enter
3. View comprehensive results including status

#### **Search Results Display**
- **Available Items**: Shows ISBN, title, Master Order ID, and status badge
- **Not Available Items**: Clear indication that item is not in repo
- **Auto-hide**: Success results disappear after 10 seconds

### **Repo Management** 🔒

#### **Access Control**
1. Click "Manage Repo" in the top navigation
2. Enter password when prompted (default: `admin123`) - characters are masked
3. Access granted for 1 hour with activity extension
4. Visual session timer shows remaining time

#### **Adding New Titles - JSON Format Only**

**JSON File Upload:**
1. Prepare JSON file with complete data.json structure
2. Upload via "Add New Titles" panel
3. Click "Process JSON File"
4. Review in preview table with status indicators

**Required JSON Format:**
```json
[
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
  }
]
```

**Status Indicators in Upload:**
- 🟢 **New (JSON)**: Valid title with complete data structure
- 🟡 **Duplicate**: Already exists in repo
- 🔴 **Invalid**: Missing required fields or malformed data

#### **Removing Titles**

**Individual Removal:**
1. Enter ISBN in "Search & Remove by ISBN"
2. Click search or press Enter
3. Found titles are automatically marked for removal

**Bulk Removal:**
1. Create file with ISBNs to remove:
   - **Text file**: One ISBN per line
   - **JSON file**: Array of objects with ISBN field
2. Upload via "Upload ISBNs to Remove"
3. Click "Process Bulk Removal"

**Removal File Formats:**
```
Text file:          JSON file:
9780140175936      [
9780140256932        {"ISBN": 9780140175936},
                     {"ISBN": 9780140256932}
                   ]
```

#### **Preview & Management**
- **New Titles Table**: Shows additions with status and individual removal
- **Removal Table**: Shows deletions with restore option
- **Statistics Dashboard**: Real-time metrics of all changes
- **Change Reversal**: Click "Restore" to undo individual removals

#### **Exporting Updated Repo**
1. **Review Changes**: Check both preview tables and statistics
2. **Backup Current**: Download current repo (recommended)
3. **Download Updated**: Get merged repo with all changes preserved
4. **Deploy**: Replace `data.json` in your GitHub repository
5. **Refresh**: Reload application to use new repo

### **Statistics Dashboard**

The repo management interface provides six key metrics:
- **Current Titles**: Existing repo count
- **To Add**: Valid new titles pending addition
- **To Remove**: Titles marked for deletion
- **Duplicates**: Duplicate entries found in uploads
- **Final Total**: Projected repo size after changes
- **Net Change**: Overall change (+/- titles)

## 🔧 Configuration

### **Changing the Password**

#### **Step 1: Generate New Hash**
Use browser console to create password hash:
```javascript
const password = "YourNewPassword";
crypto.subtle.digest('SHA-256', new TextEncoder().encode(password))
  .then(hash => {
    const hashArray = Array.from(new Uint8Array(hash));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    console.log('New password hash:', hashHex);
  });
```

#### **Step 2: Update Script**
In `script.js`, find the `INVENTORY_ACCESS` object and replace:
```javascript
const INVENTORY_ACCESS = {
    enabled: true,
    passwordHash: "your_new_hash_here",
    sessionDuration: 1 * 60 * 60 * 1000 // 1 hour
};
```

### **Session Management**
- **Duration**: Modify `sessionDuration` (in milliseconds)
- **Disable Protection**: Set `enabled: false`
- **Warning Time**: Warnings appear 10 minutes before expiration

### **Customer Information**
Update Clays POD details in `script.js`:
```javascript
const customerConfig = {
    "name": "Your Company Name",
    "address": {
        "street": "Your Street",
        "city": "Your City",
        // ... other fields
    },
    "phone": "Your Phone Number"
};
```

## 📤 CSV Output Format

Generated CSV files follow Clays POD specifications:
```csv
HDR,ORDER123,20240605,Clays POD,Clays Ltd,Popson Street,,Bungay,Suffolk,UK,NR35 1ED,GB,01986 893 211,
DTL,ORDER123,001,9781234567890,5,,,,,,,,,,
DTL,ORDER123,002,9780987654321,3,,,,,,,,,,
```

**Structure:**
- **HDR Line**: Header with order details and Clays Ltd address
- **DTL Lines**: Detail lines with line number, ISBN, and quantity

## 🛡️ Security Features

### **Content Security Policy**
- Prevents XSS attacks through resource control
- Blocks unauthorized script execution
- Validates all external resource loading

### **Input Validation**
- **File Type Checking**: Only allows specified file extensions
- **Size Limits**: 10MB maximum for all file uploads
- **ISBN Validation**: Ensures proper format before processing
- **JSON Structure**: Validates complete data.json format
- **Text Sanitization**: Removes potentially dangerous characters

### **Enhanced Session Security**
- **Masked Password Input**: Custom dialog with hidden character input
- **Encrypted Storage**: Password hashes stored securely
- **Time-based Expiration**: Automatic logout after inactivity
- **Activity Extension**: Sessions extend with user interaction
- **Manual Logout**: Immediate session termination available

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

#### **"Error loading book data"**
- **File Exists**: Confirm `data.json` is present
- **Valid JSON**: Check syntax with JSON validator
- **Proper Format**: Ensure array structure with complete book objects including Status field

#### **File Upload Problems**
- **File Size**: Maximum 10MB limit for all uploads
- **File Type**: For orders: `.xlsx`, `.xls`, `.csv`; For repo: `.json` only
- **Column Names**: Use `ISBN`/`isbn` and `Qty`/`quantity` for orders
- **JSON Structure**: Ensure complete data.json format for repo uploads
- **Order Reference**: Must be entered before file upload

#### **Password/Access Issues**
- **Correct Password**: Default is `admin123` (case-sensitive)
- **Masked Input**: Password characters are hidden as you type
- **Session Expired**: Re-authenticate if session timeout occurred
- **Browser Storage**: Ensure localStorage is enabled
- **Clear Cache**: Try clearing browser cache/data

#### **Books Showing "Not Available"**
- **ISBN Matching**: Ensure exact match with repo ISBNs
- **Format Consistency**: Check for leading zeros or formatting
- **Database Update**: Verify `data.json` contains expected entries
- **Master Order ID**: Try searching by Master Order ID if ISBN fails

#### **Status Display Issues**
- **Missing Status**: Items without Status field default to "POD Ready"
- **Status Values**: Valid values are "POD Ready", "MPI", or custom statuses
- **Badge Colors**: Green=POD Ready, Yellow=MPI, Red=Not Available

#### **Download Problems**
- **Pop-up Blockers**: Disable for the application domain
- **Download Permissions**: Check browser download settings
- **File System Access**: Ensure write permissions to download folder

### **Advanced Troubleshooting**

#### **Performance Issues**
- **Large Files**: Break large uploads into smaller batches
- **Browser Memory**: Close other tabs if processing large repos
- **Network**: Ensure stable internet for CDN resource loading

#### **Browser Compatibility**
- **Modern Browsers**: Use recent versions of Chrome, Firefox, Safari, Edge
- **Feature Support**: Verify File API and Clipboard API availability
- **JavaScript Version**: Ensure ES6+ support

#### **Security Warnings**
- **HTTPS Recommended**: Use HTTPS for production deployment
- **Mixed Content**: Ensure all resources use same protocol
- **CSP Violations**: Check console for Content Security Policy errors

## 🔄 Maintenance

### **Regular Tasks**
- **Repo Updates**: Keep `data.json` current with new titles and status changes
- **Password Rotation**: Change passwords periodically for security
- **Backup Data**: Regular backups of repo database
- **Library Updates**: Monitor CDN dependencies for updates

### **Monitoring**
- **Error Logs**: Check browser console for recurring issues
- **Usage Patterns**: Monitor which features are most used
- **Performance**: Watch for slow uploads or processing times
- **Status Accuracy**: Verify status indicators match actual production capabilities

### **Updates**
- **Dependencies**: Keep external libraries updated for security
- **Feature Requests**: Document user feedback for improvements
- **Version Control**: Use Git for tracking changes
- **Status Management**: Update status values as business requirements change

## 📊 Technical Specifications

### **Dependencies**
- **Bootstrap 5.3.2**: UI framework and responsive design
- **Font Awesome 6.4.0**: Icon library
- **SheetJS 0.18.5**: Excel file processing
- **PapaParse 5.4.1**: CSV parsing and generation

### **Browser APIs Used**
- **File API**: For file upload and processing
- **Clipboard API**: For copy-to-clipboard functionality
- **Web Crypto API**: For password hashing
- **LocalStorage API**: For session management

### **Performance Characteristics**
- **File Processing**: Up to 10MB files supported
- **Repo Size**: Tested with 10,000+ book entries
- **Session Duration**: 1-hour default with extension
- **Response Time**: Near-instantaneous for local operations

### **Data Structure Requirements**
- **Complete JSON**: All repo uploads must include full data.json structure
- **Required Fields**: ISBN, Master Order ID, TITLE, Status minimum
- **Status Values**: "POD Ready", "MPI", or custom status strings
- **ISBN Format**: 10-13 digit numbers, automatically normalized

## 🤝 Support

### **Getting Help**
- **Documentation**: Review this README for detailed guidance
- **Browser Console**: Check for error messages and warnings
- **Test Data**: Use sample data to isolate issues
- **Version Check**: Ensure all files are from the same version

### **Reporting Issues**
When reporting problems, include:
- Browser type and version
- Error messages from console
- File types and sizes being processed
- Steps to reproduce the issue
- Status values and expected behavior

### **Feature Requests**
The application is designed for extensibility. Common enhancement areas:
- Additional status types and workflows
- Enhanced security features
- Advanced reporting capabilities
- Integration with external systems
- Batch processing optimizations

## 📝 Version History

### **v4.0** - Status Field & JSON-Only Repo Management
- **Added**: Status field support with POD Ready/MPI/Not Available indicators
- **Added**: Dual search capability (ISBN and Master Order ID)
- **Added**: Enhanced filtering for MPI and Not Available items
- **Added**: JSON-only repo management with complete data structure preservation
- **Added**: Masked password input with professional dialog
- **Added**: Collapsible format documentation
- **Enhanced**: Status-aware processing and display throughout application
- **Enhanced**: Comprehensive status breakdown in loading messages
- **Simplified**: Removed Excel/CSV support from repo management
- **Improved**: Template examples with current data structure

### **v3.0** - Complete Security & Repo Management
- **Added**: Password protection for repo management
- **Added**: Session management with visual timer
- **Added**: Complete title removal functionality (individual and bulk)
- **Added**: Enhanced statistics dashboard with 6 metrics
- **Added**: Access status bar with logout functionality
- **Enhanced**: Security headers and input validation
- **Enhanced**: User interface with professional styling
- **Enhanced**: Error handling and user feedback

### **v2.0** - Enhanced Security & ISBN Search
- **Added**: Individual ISBN lookup functionality
- **Added**: Comprehensive security enhancements (CSP, input validation, XSS protection)
- **Added**: File size limits and enhanced error handling
- **Enhanced**: User interface with responsive design
- **Enhanced**: Repo addition workflow with preview tables

### **v1.0** - Initial Release
- **Core**: Excel upload, validation, and CSV export
- **Core**: Order processing with Clays POD format
- **Core**: Batch operations and clipboard integration

## 📄 License

This application is designed specifically for Clays Ltd print-on-demand operations. All configuration and branding elements are customized for Clays POD workflow requirements with status-aware processing capabilities.

---

**For technical support or customization requests, please refer to the troubleshooting section above or consult your system administrator.**