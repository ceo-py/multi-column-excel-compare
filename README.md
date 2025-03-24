
# Excel Comparison Tool

A powerful web-based tool for comparing data between two Excel files. This tool allows you to dynamically select and compare multiple columns across files, identify differences, and export the results.

![Excel Comparison Tool](https://via.placeholder.com/800x400?text=Excel+Comparison+Tool)

## Features

- **Dynamic Column Selection**: Select any columns from both files to compare
- **Multiple Comparison Mappings**: Compare multiple column pairs simultaneously
- **Flexible Key Column**: Choose which column mapping to use as the key for matching rows
- **Comprehensive Results**: View matches, differences, and rows unique to each file
- **Visual Highlighting**: Color-coded results make differences easy to spot
- **Filtering Options**: Filter results to show only what you need
- **Statistics**: See counts of matches, differences, and unique rows
- **Export to CSV**: Export your comparison results for further analysis

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/ceo-py/multi-column-excel-compare.git
   ```

2. Open `index.html` in your web browser.

3. Alternatively, you can host the files on any web server.

## Dependencies

- [SheetJS](https://sheetjs.com/) - For Excel file parsing (included via CDN)

## Usage

### 1. Load Excel Files

- Click "Choose File" to select your Excel files
- Click "Load Data" to process the files

### 2. Set Up Column Mappings

- For each column you want to compare:
  - Select a column from File 1
  - Select the corresponding column from File 2
- Add additional column mappings as needed using the "Add Column Mapping" button
- Choose which column mapping should be used as the key for matching rows

### 3. Compare Files

- Click "Compare Files" to analyze the data
- View the results in the table below
- Use the checkboxes to filter which types of results to display

### 4. Export Results

- Click "Export to CSV" to download the comparison results

## How It Works

1. **File Loading**: The tool reads Excel files using the SheetJS library and extracts column headers and data.

2. **Column Mapping**: You define which columns should be compared between the two files.

3. **Key Column**: One column mapping is designated as the "key" to match rows between files.

4. **Comparison Process**:
   - Rows are matched between files using the key column values
   - For each matched pair of rows, all mapped columns are compared
   - Rows are categorized as: matches, differences, unique to file 1, or unique to file 2

5. **Results Display**:
   - Results are shown in a table with color coding
   - Statistics show counts of matches, differences, and unique rows
   - Filtering options let you focus on specific result types

## Example

Let's say you have two Excel files with customer data:

**File 1 (Current Customers):**
| Customer ID | Name | Email | Plan | Monthly Fee |
|-------------|------|-------|------|-------------|
| 1001 | John Smith | john@example.com | Premium | 29.99 |
| 1002 | Jane Doe | jane@example.com | Basic | 9.99 |
| 1003 | Bob Johnson | bob@example.com | Standard | 19.99 |

**File 2 (Updated Customers):**
| ID | Customer Name | Email Address | Subscription | Fee |
|-------------|------|-------|------|-------------|
| 1001 | John Smith | john@example.com | Premium | 29.99 |
| 1002 | Jane Doe | jane@example.com | Premium | 29.99 |
| 1004 | Sarah Wilson | sarah@example.com | Basic | 9.99 |

You can set up the following column mappings:
1. "Customer ID" in File 1 maps to "ID" in File 2 (use as key)
2. "Name" in File 1 maps to "Customer Name" in File 2
3. "Email" in File 1 maps to "Email Address" in File 2
4. "Plan" in File 1 maps to "Subscription" in File 2
5. "Monthly Fee" in File 1 maps to "Fee" in File 2

The comparison results would show:
- Customer 1001: Match (all values are the same)
- Customer 1002: Different (Plan/Subscription and Monthly Fee/Fee values differ)
- Customer 1003: Unique to File 1
- Customer 1004: Unique to File 2

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
