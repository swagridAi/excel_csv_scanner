# report/templates.py
"""
Templates for report generation.
"""

EXCEL_TEMPLATE_HEADER = [
    "Filename", 
    "Extension", 
    "Size (KB)", 
    "Created Date",
    "Modified Date", 
    "Accessed Date",
    "Creator",
    "Last Modified By",
    "Is Excel", 
    "Has VBA", 
    "Has PivotTable", 
    "Row Count", 
    "Column Count",
    "Sheet Count",
    "Delimiter",
    "Encoding"
]

# HTML report template with enhanced support for user and timestamp information
HTML_REPORT_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Excel and CSV File Analysis Report</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 2px solid #eaecef;
            padding-bottom: 10px;
        }
        h2 {
            color: #3498db;
            margin-top: 30px;
        }
        .container {
            background-color: #fff;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            padding: 20px;
            margin-bottom: 20px;
        }
        .summary {
            margin-bottom: 30px;
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            border-left: 5px solid #3498db;
        }
        .user-summary {
            margin-bottom: 30px;
            background-color: #e8f4fd;
            padding: 15px;
            border-radius: 5px;
            border-left: 5px solid #2980b9;
        }
        .timeline-summary {
            margin-bottom: 30px;
            background-color: #fdeaea;
            padding: 15px;
            border-radius: 5px;
            border-left: 5px solid #e74c3c;
        }
        .meta-info {
            color: #7f8c8d;
            font-size: 0.9em;
            margin-bottom: 20px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 20px;
            font-size: 0.9em;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            position: sticky;
            top: 0;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .data-table-container {
            max-height: 700px;
            overflow-y: auto;
            margin-top: 20px;
        }
        .summary-table {
            width: auto;
            min-width: 300px;
        }
        .summary-table th {
            background-color: #3498db;
            color: white;
        }
        .error {
            color: #e74c3c;
        }
        .footer {
            margin-top: 30px;
            padding-top: 10px;
            border-top: 1px solid #eaecef;
            color: #7f8c8d;
            font-size: 0.8em;
            text-align: center;
        }
        
        /* Enhanced styles for user and timestamp metadata */
        .user-column {
            background-color: #D4E6F1 !important;
            color: #2471A3;
            font-weight: bold;
        }
        .date-column {
            background-color: #FADBD8 !important;
            color: #943126;
            font-weight: bold;
        }
        .user-cell {
            background-color: #EBF5FB !important;
        }
        .date-cell {
            background-color: #FDEDEC !important;
        }
        
        /* Tab styles */
        .tabs {
            display: flex;
            border-bottom: 1px solid #ddd;
            margin-bottom: 20px;
        }
        .tab {
            padding: 10px 20px;
            cursor: pointer;
            border: 1px solid transparent;
            border-bottom: none;
            border-radius: 5px 5px 0 0;
            margin-right: 5px;
            background-color: #f1f1f1;
        }
        .tab:hover {
            background-color: #ddd;
        }
        .tab.active {
            background-color: white;
            border-color: #ddd;
            border-bottom: 1px solid white;
            margin-bottom: -1px;
            font-weight: bold;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
        }
        
        /* Cards for user info */
        .user-cards {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin-top: 20px;
        }
        .user-card {
            background-color: #EBF5FB;
            border-radius: 5px;
            padding: 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            width: calc(33.33% - 20px);
            min-width: 250px;
        }
        .user-card h3 {
            margin-top: 0;
            color: #2980b9;
        }
        .user-card p {
            margin: 5px 0;
        }
        .user-card .file-info {
            font-size: 0.8em;
            color: #7f8c8d;
            margin-top: 10px;
            padding-top: 10px;
            border-top: 1px solid #d4e6f1;
        }
        
        /* Timeline visualization */
        .timeline {
            margin-top: 20px;
            position: relative;
            padding-left: 30px;
        }
        .timeline::before {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            height: 100%;
            width: 2px;
            background-color: #3498db;
        }
        .timeline-item {
            margin-bottom: 20px;
            position: relative;
        }
        .timeline-item::before {
            content: '';
            position: absolute;
            left: -25px;
            top: 5px;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background-color: #3498db;
        }
        .timeline-item.created::before {
            background-color: #2ecc71;
        }
        .timeline-item.modified::before {
            background-color: #f39c12;
        }
        .timeline-item.accessed::before {
            background-color: #9b59b6;
        }
        .timeline-item .time {
            font-weight: bold;
            color: #2c3e50;
        }
        .timeline-item .event {
            font-weight: bold;
        }
        .timeline-item .file {
            font-style: italic;
        }
        .timeline-item .user {
            color: #2980b9;
        }
        
        /* Additional styling can be added here */
        {additional_css}
    </style>
    <script>
        // JavaScript to enable tabbed interface
        document.addEventListener('DOMContentLoaded', function() {
            // Set up tab functionality
            const tabs = document.querySelectorAll('.tab');
            tabs.forEach(tab => {
                tab.addEventListener('click', function() {
                    // Remove active class from all tabs and content
                    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                    
                    // Add active class to clicked tab and corresponding content
                    this.classList.add('active');
                    const target = this.getAttribute('data-target');
                    document.getElementById(target).classList.add('active');
                });
            });
            
            // Activate first tab by default
            if (tabs.length > 0) {
                tabs[0].click();
            }
            
            // Set up sorting for tables
            document.querySelectorAll('th').forEach(th => {
                th.addEventListener('click', function() {
                    const table = this.closest('table');
                    const rows = Array.from(table.querySelectorAll('tbody tr'));
                    const columnIndex = Array.from(this.parentNode.children).indexOf(this);
                    
                    // Toggle sort direction
                    const sortDir = this.getAttribute('data-sort') === 'asc' ? 'desc' : 'asc';
                    
                    // Reset sort indicators on all columns
                    table.querySelectorAll('th').forEach(header => {
                        header.removeAttribute('data-sort');
                    });
                    
                    // Set sort indicator on current column
                    this.setAttribute('data-sort', sortDir);
                    
                    // Sort rows
                    rows.sort((a, b) => {
                        const aValue = a.children[columnIndex].textContent.trim();
                        const bValue = b.children[columnIndex].textContent.trim();
                        
                        // Try numeric sort if possible
                        const aNum = parseFloat(aValue);
                        const bNum = parseFloat(bValue);
                        
                        if (!isNaN(aNum) && !isNaN(bNum)) {
                            return sortDir === 'asc' ? aNum - bNum : bNum - aNum;
                        }
                        
                        // Fallback to string sort
                        if (sortDir === 'asc') {
                            return aValue.localeCompare(bValue);
                        } else {
                            return bValue.localeCompare(aValue);
                        }
                    });
                    
                    // Reorder rows in the table
                    const tbody = table.querySelector('tbody');
                    rows.forEach(row => tbody.appendChild(row));
                });
            });
        });
    </script>
</head>
<body>
    <h1>Excel and CSV File Analysis Report</h1>
    
    <div class="meta-info">
        <p>Generated on: {date}</p>
    </div>
    
    <!-- Summary information -->
    {summary}
    
    <!-- Tabbed interface -->
    <div class="container">
        <div class="tabs">
            <div class="tab active" data-target="tab-files">File Analysis</div>
            <div class="tab" data-target="tab-users">User Information</div>
            <div class="tab" data-target="tab-timeline">Timeline</div>
        </div>
        
        <!-- File Analysis Tab -->
        <div id="tab-files" class="tab-content active">
            <h2>File Analysis Results</h2>
            <div class="data-table-container">
                <table class="data-table">
                    <thead>
                        <tr>
                            {headers}
                        </tr>
                    </thead>
                    <tbody>
                        {rows}
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- User Information Tab -->
        <div id="tab-users" class="tab-content">
            <h2>User Information</h2>
            <p>This tab shows information about users who have created or modified files in the dataset.</p>
            
            <div class="user-cards">
                {user_cards}
            </div>
        </div>
        
        <!-- Timeline Tab -->
        <div id="tab-timeline" class="tab-content">
            <h2>File Timeline</h2>
            <p>This tab shows a chronological timeline of file creation, modification, and access events.</p>
            
            <div class="timeline">
                {timeline_items}
            </div>
        </div>
    </div>
    
    <div class="footer">
        <p>Generated by Excel and CSV Inspector Tool</p>
    </div>
</body>
</html>
"""

# Basic text report template
TEXT_REPORT_TEMPLATE = """
Excel and CSV File Analysis Report
=================================
Generated on: {date}

Summary
-------
Total files analyzed: {total_files}
Excel files: {excel_files}
CSV files: {csv_files}
Files with VBA: {vba_files}
Files with PivotTables: {pivot_files}

User Information
---------------
Files with user metadata: {files_with_user_info}
Unique creators identified: {unique_creators}
Unique editors identified: {unique_editors}

Timestamp Information
-------------------
Files with creation dates: {files_with_creation_dates}
Files with modification dates: {files_with_modification_dates}
Files with access dates: {files_with_access_dates}

Detailed Results
---------------
{results}

"""

# JSON report schema template
JSON_SCHEMA = {
    "metadata": {
        "generated_at": "",  # ISO timestamp
        "file_count": 0,
        "excel_count": 0,
        "csv_count": 0,
        "vba_count": 0,
        "pivot_count": 0,
        "files_with_user_info": 0,
        "files_with_complete_metadata": 0,
        "unique_users": []
    },
    "results": []  # List of file analysis results
}