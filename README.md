MCP Excel Server
A Model Context Protocol (MCP) server that provides Excel file operations including reading, writing, formatting, and chart creation.

Features
Read Excel Files: Extract data from existing Excel workbooks and worksheets
Write Excel Files: Create new workbooks and write data to specific cells or ranges
Format Cells: Apply font styles, colors, and background colors to cell ranges
Create Charts: Generate bar charts, line charts, and pie charts from data
Add Formulas: Insert Excel formulas into cells for calculations
Multiple Sheet Support: Work with multiple worksheets within a workbook
Installation
Clone or create the project directory:
bash
mkdir mcp-excel-server
cd mcp-excel-server
Create a virtual environment:
bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
Install dependencies:
bash
pip install -r requirements.txt
Usage
Running the Server
Start the MCP server:

bash
python excel_server.py
Available Tools
1. read_excel
Read data from an Excel file.

Parameters:

file_path: Path to the Excel file
sheet_name: Name of the worksheet to read from
range (optional): Specific cell range to read (e.g., "A1:C10")
2. write_excel
Write data to an Excel file.

Parameters:

file_path: Path to the Excel file
sheet_name: Name of the worksheet to write to
data: 2D array of data to write
start_cell (optional): Starting cell position (default: "A1")
3. create_workbook
Create a new Excel workbook.

Parameters:

file_path: Path for the new Excel file
sheet_names (optional): List of sheet names to create (default: ["Sheet1"])
4. format_cells
Apply formatting to Excel cells.

Parameters:

file_path: Path to the Excel file
sheet_name: Name of the worksheet
range: Cell range to format
font_bold (optional): Make text bold
font_size (optional): Font size
bg_color (optional): Background color (hex format)
font_color (optional): Font color (hex format)
5. create_chart
Create a chart in Excel.

Parameters:

file_path: Path to the Excel file
sheet_name: Name of the worksheet
chart_type: Type of chart ("bar", "line", "pie")
data_range: Data range for the chart
title (optional): Chart title
position (optional): Chart position cell (default: "E5")
6. add_formula
Add a formula to a cell.

Parameters:

file_path: Path to the Excel file
sheet_name: Name of the worksheet
cell: Target cell (e.g., "A1")
formula: Excel formula (without the leading "=")
Testing
Run the test client to verify the server works correctly:

bash
python test_client.py
This will:

Create a new workbook with sample data
Add formulas for calculations
Format the header row
Create a bar chart
Read the data back to verify
Example Usage
Here's a simple example of using the MCP Excel server:

python
# Create a new workbook
await session.call_tool("create_workbook", {
    "file_path": "sales_report.xlsx",
    "sheet_names": ["Q1_Sales", "Q2_Sales"]
})

# Write sales data
sales_data = [
    ["Product", "Jan", "Feb", "Mar", "Total"],
    ["Product A", 1000, 1200, 1100, "=SUM(B2:D2)"],
    ["Product B", 800, 900, 950, "=SUM(B3:D3)"],
    ["Product C", 1500, 1400, 1600, "=SUM(B4:D4)"]
]

await session.call_tool("write_excel", {
    "file_path": "sales_report.xlsx",
    "sheet_name": "Q1_Sales",
    "data": sales_data
})

# Format the header row
await session.call_tool("format_cells", {
    "file_path": "sales_report.xlsx",
    "sheet_name": "Q1_Sales",
    "range": "A1:E1",
    "font_bold": True,
    "bg_color": "#366092",
    "font_color": "#FFFFFF"
})

# Create a chart
await session.call_tool("create_chart", {
    "file_path": "sales_report.xlsx",
    "sheet_name": "Q1_Sales",
    "chart_type": "bar",
    "data_range": "A1:D4",
    "title": "Q1 Sales by Product",
    "position": "G2"
})
Requirements
Python 3.8+
mcp>=1.0.0
openpyxl>=3.1.0
License
MIT License - see LICENSE file for details.

Contributing
Fork the repository
Create a feature branch
Make your changes
Add tests for new functionality
Submit a pull request
Error Handling
The server includes comprehensive error handling for:

File not found errors
Invalid sheet names
Invalid cell ranges
Unsupported chart types
Formula syntax errors
All errors are returned as JSON responses with descriptive error messages.

