// Global variables to store the loaded data
let file1Data = null;
let file2Data = null;
let file1Columns = [];
let file2Columns = [];
let comparisonResults = [];

// DOM elements
const loadDataBtn = document.getElementById('load-data-btn');
const comparisonSection = document.getElementById('comparison-section');
const resultsSection = document.getElementById('results-section');
const columnMappingsContainer = document.getElementById('column-mappings');
const addMappingBtn = document.getElementById('add-mapping-btn');
const compareBtn = document.getElementById('compare-btn');
const keyColumnSelect = document.getElementById('key-column-select');
const resultsTable = document.getElementById('results-table');
const resultsHeader = document.getElementById('results-header');
const resultsBody = document.getElementById('results-body');
const exportBtn = document.getElementById('export-btn');

// Filter checkboxes
const showMatches = document.getElementById('show-matches');
const showDifferences = document.getElementById('show-differences');
const showUniqueFile1 = document.getElementById('show-unique-file1');
const showUniqueFile2 = document.getElementById('show-unique-file2');

// Statistics elements
const totalRowsElement = document.querySelector('#total-rows span');
const matchingRowsElement = document.querySelector('#matching-rows span');
const differentRowsElement = document.querySelector('#different-rows span');
const uniqueFile1Element = document.querySelector('#unique-file1 span');
const uniqueFile2Element = document.querySelector('#unique-file2 span');

// Event listeners
loadDataBtn.addEventListener('click', loadExcelFiles);
addMappingBtn.addEventListener('click', addColumnMapping);
compareBtn.addEventListener('click', compareFiles);
exportBtn.addEventListener('click', exportResults);

// Filter change events
showMatches.addEventListener('change', filterResults);
showDifferences.addEventListener('change', filterResults);
showUniqueFile1.addEventListener('change', filterResults);
showUniqueFile2.addEventListener('change', filterResults);

/**
 * Load and parse Excel files
 */
async function loadExcelFiles() {
    const file1Input = document.getElementById('file1');
    const file2Input = document.getElementById('file2');
    
    if (!file1Input.files.length || !file2Input.files.length) {
        alert('Please select both files');
        return;
    }
    
    try {
        // Load file 1
        file1Data = await readExcelFile(file1Input.files[0]);
        file1Columns = Object.keys(file1Data[0] || {});
        
        // Load file 2
        file2Data = await readExcelFile(file2Input.files[0]);
        file2Columns = Object.keys(file2Data[0] || {});
        
        if (file1Columns.length === 0 || file2Columns.length === 0) {
            alert('One or both files appear to be empty');
            return;
        }
        
        // Show comparison section and initialize column selects
        comparisonSection.style.display = 'block';
        initializeColumnMappings();
        
    } catch (error) {
        console.error('Error loading files:', error);
        alert('Error loading files: ' + error.message);
    }
}

/**
 * Read and parse Excel file using SheetJS
 */
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Get first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 'A' });
                
                // If the first row contains headers, remove it and use as column names
                if (jsonData.length > 0) {
                    const headers = jsonData.shift();
                    
                    // Convert remaining data to use header names as keys
                    const processedData = jsonData.map(row => {
                        const newRow = {};
                        Object.keys(row).forEach(key => {
                            newRow[headers[key] || key] = row[key];
                        });
                        return newRow;
                    });
                    
                    resolve(processedData);
                } else {
                    resolve([]);
                }
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = function(error) {
            reject(error);
        };
        
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Initialize column mapping UI with the first mapping row
 */
function initializeColumnMappings() {
    // Clear existing mappings
    columnMappingsContainer.innerHTML = '';
    
    // Add the first mapping row
    addColumnMapping();
    
    // Reset key column select
    keyColumnSelect.innerHTML = '<option value="0">First Column Mapping</option>';
}

/**
 * Add a new column mapping row to the UI
 */
function addColumnMapping() {
    const mappingRow = document.createElement('div');
    mappingRow.className = 'column-mapping-row';
    
    // Create column 1 select
    const column1Container = document.createElement('div');
    column1Container.className = 'column-select';
    
    const column1Label = document.createElement('label');
    column1Label.textContent = 'File 1 Column:';
    
    const column1Select = document.createElement('select');
    column1Select.className = 'column1-select';
    
    file1Columns.forEach((column, index) => {
        const option = document.createElement('option');
        option.value = column;
        option.textContent = column;
        column1Select.appendChild(option);
    });
    
    column1Container.appendChild(column1Label);
    column1Container.appendChild(column1Select);
    
    // Create comparison operator
    const operatorContainer = document.createElement('div');
    operatorContainer.className = 'comparison-operator';
    operatorContainer.innerHTML = '<span>compared to</span>';
    
    // Create column 2 select
    const column2Container = document.createElement('div');
    column2Container.className = 'column-select';
    
    const column2Label = document.createElement('label');
    column2Label.textContent = 'File 2 Column:';
    
    const column2Select = document.createElement('select');
    column2Select.className = 'column2-select';
    
    file2Columns.forEach((column, index) => {
        const option = document.createElement('option');
        option.value = column;
        option.textContent = column;
        column2Select.appendChild(option);
    });
    
    column2Container.appendChild(column2Label);
    column2Container.appendChild(column2Select);
    
    // Create remove button
    const removeBtn = document.createElement('button');
    removeBtn.className = 'remove-mapping-btn';
    removeBtn.title = 'Remove this mapping';
    removeBtn.textContent = '×';
    removeBtn.addEventListener('click', function() {
        // Don't allow removing if it's the last mapping
        if (columnMappingsContainer.children.length > 1) {
            mappingRow.remove();
            updateKeyColumnOptions();
        } else {
            alert('You need at least one column mapping');
        }
    });
    
    // Assemble the row
    mappingRow.appendChild(column1Container);
    mappingRow.appendChild(operatorContainer);
    mappingRow.appendChild(column2Container);
    mappingRow.appendChild(removeBtn);
    
    // Add to container
    columnMappingsContainer.appendChild(mappingRow);
    
    // Update key column options
    updateKeyColumnOptions();
}

/**
 * Update the key column dropdown options
 */
function updateKeyColumnOptions() {
    keyColumnSelect.innerHTML = '';
    
    const mappingRows = columnMappingsContainer.querySelectorAll('.column-mapping-row');
    mappingRows.forEach((row, index) => {
        const column1Select = row.querySelector('.column1-select');
        const column2Select = row.querySelector('.column2-select');
        
        const option = document.createElement('option');
        option.value = index;
        option.textContent = `${column1Select.value} ↔ ${column2Select.value}`;
        keyColumnSelect.appendChild(option);
    });
}

/**
 * Compare the files based on the selected column mappings
 */
function compareFiles() {
    // Get all column mappings
    const mappingRows = columnMappingsContainer.querySelectorAll('.column-mapping-row');
    const columnMappings = Array.from(mappingRows).map(row => {
        const column1 = row.querySelector('.column1-select').value;
        const column2 = row.querySelector('.column2-select').value;
        return { column1, column2 };
    });
    
    if (columnMappings.length === 0) {
        alert('Please add at least one column mapping');
        return;
    }
    
    // Get key column mapping
    const keyMappingIndex = parseInt(keyColumnSelect.value);
    const keyMapping = columnMappings[keyMappingIndex];
    
    // Create lookup maps for both files based on key column
    const file1Map = new Map();
    file1Data.forEach((row, index) => {
        const keyValue = row[keyMapping.column1];
        if (keyValue !== undefined) {
            file1Map.set(keyValue.toString(), { data: row, index });
        }
    });
    
    const file2Map = new Map();
    file2Data.forEach((row, index) => {
        const keyValue = row[keyMapping.column2];
        if (keyValue !== undefined) {
            file2Map.set(keyValue.toString(), { data: row, index });
        }
    });
    
    // Compare data
    comparisonResults = [];
    let matchCount = 0;
    let diffCount = 0;
    let uniqueFile1Count = 0;
    let uniqueFile2Count = 0;
    
    // Check rows in file 1
    file1Map.forEach((file1Row, key) => {
        if (file2Map.has(key)) {
            // Row exists in both files
            const file2Row = file2Map.get(key);
            
            // Check if all mapped columns match
            let allMatch = true;
            const differences = [];
            
            columnMappings.forEach(mapping => {
                const value1 = file1Row.data[mapping.column1];
                const value2 = file2Row.data[mapping.column2];
                
                // Convert to strings for comparison
                const strValue1 = value1 !== undefined ? value1.toString() : '';
                const strValue2 = value2 !== undefined ? value2.toString() : '';
                
                if (strValue1 !== strValue2) {
                    allMatch = false;
                    differences.push({
                        column1: mapping.column1,
                        column2: mapping.column2,
                        value1: strValue1,
                        value2: strValue2
                    });
                }
            });
            
            if (allMatch) {
                comparisonResults.push({
                    key,
                    file1Index: file1Row.index,
                    file2Index: file2Row.index,
                    file1Data: file1Row.data,
                    file2Data: file2Row.data,
                    status: 'match',
                    differences: []
                });
                matchCount++;
            } else {
                comparisonResults.push({
                    key,
                    file1Index: file1Row.index,
                    file2Index: file2Row.index,
                    file1Data: file1Row.data,
                    file2Data: file2Row.data,
                    status: 'different',
                    differences
                });
                diffCount++;
            }
        } else {
            // Row only exists in file 1
            comparisonResults.push({
                key,
                file1Index: file1Row.index,
                file2Index: null,
                file1Data: file1Row.data,
                file2Data: null,
                status: 'unique-file1',
                differences: []
            });
            uniqueFile1Count++;
        }
    });
    
    // Check for rows only in file 2
    file2Map.forEach((file2Row, key) => {
        if (!file1Map.has(key)) {
            comparisonResults.push({
                key,
                file1Index: null,
                file2Index: file2Row.index,
                file1Data: null,
                file2Data: file2Row.data,
                status: 'unique-file2',
                differences: []
            });
            uniqueFile2Count++;
        }
    });
    
    // Update statistics
    totalRowsElement.textContent = comparisonResults.length;
    matchingRowsElement.textContent = matchCount;
    differentRowsElement.textContent = diffCount;
    uniqueFile1Element.textContent = uniqueFile1Count;
    uniqueFile2Element.textContent = uniqueFile2Count;
    
    // Display results
    displayResults(columnMappings);
    resultsSection.style.display = 'block';
}

/**
 * Display comparison results in the table
 */
function displayResults(columnMappings) {
    // Create table header
    const headerRow = resultsHeader.querySelector('tr');
    headerRow.innerHTML = '<th>Row #</th>';
    
    // Add column headers for each mapping
    columnMappings.forEach(mapping => {
        const th1 = document.createElement('th');
        th1.textContent = `${mapping.column1} (File 1)`;
        headerRow.appendChild(th1);
        
        const th2 = document.createElement('th');
        th2.textContent = `${mapping.column2} (File 2)`;
        headerRow.appendChild(th2);
    });
    
    headerRow.appendChild(document.createElement('th')).textContent = 'Status';
    
    // Apply filters and display results
    filterResults();
}

/**
 * Filter and display results based on checkbox selections
 */
function filterResults() {
    resultsBody.innerHTML = '';
    
    // Get filter states
    const showMatchesChecked = showMatches.checked;
    const showDifferencesChecked = showDifferences.checked;
    const showUniqueFile1Checked = showUniqueFile1.checked;
    const showUniqueFile2Checked = showUniqueFile2.checked;
    
    // Get column mappings
    const mappingRows = columnMappingsContainer.querySelectorAll('.column-mapping-row');
    const columnMappings = Array.from(mappingRows).map(row => {
        const column1 = row.querySelector('.column1-select').value;
        const column2 = row.querySelector('.column2-select').value;
        return { column1, column2 };
    });
    
    // Filter results
    const filteredResults = comparisonResults.filter(result => {
        switch (result.status) {
            case 'match': return showMatchesChecked;
            case 'different': return showDifferencesChecked;
            case 'unique-file1': return showUniqueFile1Checked;
            case 'unique-file2': return showUniqueFile2Checked;
            default: return false;
        }
    });
    
    // Display filtered results
    filteredResults.forEach((result, index) => {
        const row = document.createElement('tr');
        row.className = result.status;
        
        // Row number
        const rowNumCell = document.createElement('td');
        rowNumCell.textContent = index + 1;
        row.appendChild(rowNumCell);
        
        // Data cells for each column mapping
        columnMappings.forEach(mapping => {
            const file1Value = result.file1Data ? result.file1Data[mapping.column1] || '' : '';
            const file2Value = result.file2Data ? result.file2Data[mapping.column2] || '' : '';
            
            // Check if this column has a difference
            const hasDifference = result.differences.some(diff => 
                diff.column1 === mapping.column1 && diff.column2 === mapping.column2
            );
            
            // File 1 value cell
            const cell1 = document.createElement('td');
            cell1.textContent = file1Value;
            if (hasDifference) cell1.className = 'different-value';
            row.appendChild(cell1);
            
            // File 2 value cell
            const cell2 = document.createElement('td');
            cell2.textContent = file2Value;
            if (hasDifference) cell2.className = 'different-value';
            row.appendChild(cell2);
        });
        
        // Status cell
        const statusCell = document.createElement('td');
        statusCell.className = 'status';
        
        switch (result.status) {
            case 'match':
                statusCell.textContent = 'Match';
                statusCell.classList.add('match');
                break;
            case 'different':
                statusCell.textContent = 'Different';
                statusCell.classList.add('different');
                break;
            case 'unique-file1':
                statusCell.textContent = 'Only in File 1';
                statusCell.classList.add('unique-file1');
                break;
            case 'unique-file2':
                statusCell.textContent = 'Only in File 2';
                statusCell.classList.add('unique-file2');
                break;
        }
        
        row.appendChild(statusCell);
        resultsBody.appendChild(row);
    });
}

/**
 * Export results to CSV file
 */
function exportResults() {
    if (comparisonResults.length === 0) {
        alert('No results to export');
        return;
    }
    
    // Get column mappings
    const mappingRows = columnMappingsContainer.querySelectorAll('.column-mapping-row');
    const columnMappings = Array.from(mappingRows).map(row => {
        const column1 = row.querySelector('.column1-select').value;
        const column2 = row.querySelector('.column2-select').value;
        return { column1, column2 };
    });
    
    // Create CSV header
    let csvContent = 'Row,';
    
    // Add column headers
    columnMappings.forEach(mapping => {
        csvContent += `"${mapping.column1} (File 1)",`;
        csvContent += `"${mapping.column2} (File 2)",`;
    });
    
    csvContent += 'Status,Differences\n';
    
    // Add data rows
    comparisonResults.forEach((result, index) => {
        csvContent += `${index + 1},`;
        
        // Add data for each column mapping
        columnMappings.forEach(mapping => {
            const file1Value = result.file1Data ? result.file1Data[mapping.column1] || '' : '';
            const file2Value = result.file2Data ? result.file2Data[mapping.column2] || '' : '';
            
            csvContent += `"${file1Value}",`;
            csvContent += `"${file2Value}",`;
        });
        
        // Add status
        let status = '';
        switch (result.status) {
            case 'match': status = 'Match'; break;
            case 'different': status = 'Different'; break;
            case 'unique-file1': status = 'Only in File 1'; break;
            case 'unique-file2': status = 'Only in File 2'; break;
        }
        csvContent += `${status},`;
        
        // Add differences description
        if (result.differences.length > 0) {
            const diffText = result.differences.map(diff => 
                `${diff.column1}:${diff.value1} ≠ ${diff.column2}:${diff.value2}`
            ).join('; ');
            csvContent += `"${diffText}"`;
        }
        
        csvContent += '\n';
    });
    
    // Create and download CSV file
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'excel_comparison_results.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}