import "./taskpane.css";
import { 
  convertToDatatable, 
  inferColumns, 
  ColumnInfo, 
  KustoDataType,
  KUSTO_DATA_TYPES 
} from "../services/datatableConverter";

// Store state for the current conversion
let currentData: unknown[][] = [];
let currentColumns: ColumnInfo[] = [];
let lastConvertedResult: string = "";

// Initialize when Office.js is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    updateStatus("Ready. Select cells and click Load Selection.");
    initializeUI();
  } else {
    updateStatus("This add-in only works in Excel.");
  }
});

/**
 * Sets up UI event handlers after Office.js is ready
 */
function initializeUI(): void {
  const loadBtn = document.getElementById("load-btn") as HTMLButtonElement;
  const convertBtn = document.getElementById("convert-btn") as HTMLButtonElement;
  const copyBtn = document.getElementById("copy-btn") as HTMLButtonElement;
  const headerCheckbox = document.getElementById("first-row-header") as HTMLInputElement;

  if (loadBtn) {
    loadBtn.disabled = false;
    loadBtn.addEventListener("click", handleLoadClick);
  }

  if (convertBtn) {
    convertBtn.addEventListener("click", handleConvertClick);
  }

  if (copyBtn) {
    copyBtn.addEventListener("click", handleCopyClick);
  }

  if (headerCheckbox) {
    headerCheckbox.addEventListener("change", handleHeaderCheckboxChange);
  }
}

/**
 * Handles the Load Selection button click - reads the selected range and shows column mappings
 */
async function handleLoadClick(): Promise<void> {
  const loadBtn = document.getElementById("load-btn") as HTMLButtonElement;
  const headerCheckbox = document.getElementById("first-row-header") as HTMLInputElement;
  
  try {
    loadBtn.disabled = true;
    updateStatus("Loading selection...");
    hideOutput();

    currentData = await readSelectedRange();
    const firstRowIsHeader = headerCheckbox?.checked ?? true;
    
    // Validate we have data
    if (!currentData || currentData.length === 0) {
      throw new Error("No data to convert. Please select cells with data.");
    }

    const minRows = firstRowIsHeader ? 2 : 1;
    if (currentData.length < minRows) {
      throw new Error(firstRowIsHeader 
        ? "No data rows to convert. If first row is header, select at least 2 rows."
        : "No data to convert. Please select cells with data.");
    }

    // Infer column types
    currentColumns = inferColumns(currentData, firstRowIsHeader);
    
    // Display the column mapping UI
    displayColumnMappings(currentColumns);
    
    const rowCount = firstRowIsHeader ? currentData.length - 1 : currentData.length;
    updateStatus(`Loaded ${currentColumns.length} column(s) and ${rowCount} data row(s). Adjust types and click Convert.`);

  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown error occurred";
    updateStatus(`Error: ${message}`);
    hideColumnMappings();
    currentData = [];
    currentColumns = [];
  } finally {
    loadBtn.disabled = false;
  }
}

/**
 * Displays the column type mapping UI
 */
function displayColumnMappings(columns: ColumnInfo[]): void {
  const section = document.getElementById("column-mapping-section");
  const container = document.getElementById("column-mappings");
  
  if (!section || !container) return;

  // Clear existing mappings
  container.innerHTML = "";

  // Create a row for each column
  columns.forEach((col, index) => {
    const row = document.createElement("div");
    row.className = "column-mapping-row";
    
    const nameSpan = document.createElement("span");
    nameSpan.className = "column-name";
    nameSpan.textContent = col.name;
    nameSpan.title = col.name; // Tooltip for long names
    
    const typeSelect = document.createElement("select");
    typeSelect.id = `column-type-${index}`;
    typeSelect.dataset.columnIndex = index.toString();
    
    // Add all Kusto types as options
    KUSTO_DATA_TYPES.forEach(type => {
      const option = document.createElement("option");
      option.value = type;
      option.textContent = type;
      if (type === col.type) {
        option.selected = true;
      }
      typeSelect.appendChild(option);
    });

    // Update the stored column type when changed
    typeSelect.addEventListener("change", (e) => {
      const target = e.target as HTMLSelectElement;
      const colIndex = parseInt(target.dataset.columnIndex!, 10);
      currentColumns[colIndex].type = target.value as KustoDataType;
    });
    
    row.appendChild(nameSpan);
    row.appendChild(typeSelect);
    container.appendChild(row);
  });

  section.classList.remove("hidden");
}

/**
 * Hides the column mapping UI
 */
function hideColumnMappings(): void {
  const section = document.getElementById("column-mapping-section");
  if (section) {
    section.classList.add("hidden");
  }
}

/**
 * Handles the Convert button click - uses the user-specified column types
 */
async function handleConvertClick(): Promise<void> {
  const headerCheckbox = document.getElementById("first-row-header") as HTMLInputElement;
  
  try {
    updateStatus("Converting...");
    
    const firstRowIsHeader = headerCheckbox?.checked ?? true;
    
    // Build column types map from user selections
    const columnTypes = new Map<number, KustoDataType>();
    currentColumns.forEach((col, index) => {
      columnTypes.set(index, col.type);
    });
    
    // Convert to Kusto datatable syntax
    lastConvertedResult = convertToDatatable(currentData, { 
      firstRowIsHeader,
      columnTypes 
    });
    
    displayOutput(lastConvertedResult);
    
    const rowCount = firstRowIsHeader ? currentData.length - 1 : currentData.length;
    updateStatus(`Converted ${rowCount} data row(s) to datatable.`);

  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown error occurred";
    updateStatus(`Error: ${message}`);
    hideOutput();
    lastConvertedResult = "";
  }
}

/**
 * Handles the Copy to Clipboard button click
 */
async function handleCopyClick(): Promise<void> {
  if (!lastConvertedResult) {
    updateStatus("Nothing to copy. Convert a selection first.");
    return;
  }

  try {
    await navigator.clipboard.writeText(lastConvertedResult);
    showCopyConfirmation();
    updateStatus("Copied to clipboard!");
  } catch (error) {
    // Fallback for browsers that don't support clipboard API
    fallbackCopyToClipboard(lastConvertedResult);
  }
}

/**
 * Fallback copy method using a temporary textarea
 */
function fallbackCopyToClipboard(text: string): void {
  const textArea = document.createElement("textarea");
  textArea.value = text;
  textArea.style.position = "fixed";
  textArea.style.left = "-9999px";
  document.body.appendChild(textArea);
  textArea.select();
  
  try {
    document.execCommand("copy");
    showCopyConfirmation();
    updateStatus("Copied to clipboard!");
  } catch (error) {
    updateStatus("Failed to copy. Please select and copy manually.");
  } finally {
    document.body.removeChild(textArea);
  }
}

/**
 * Shows the copy confirmation message briefly
 */
function showCopyConfirmation(): void {
  const confirmation = document.getElementById("copy-confirmation");
  if (confirmation) {
    confirmation.classList.remove("hidden");
    setTimeout(() => {
      confirmation.classList.add("hidden");
    }, 2000);
  }
}

/**
 * Handles the header checkbox change - reload if we have data
 */
async function handleHeaderCheckboxChange(): Promise<void> {
  // If we have data loaded, re-infer the columns with the new header setting
  if (currentData.length > 0) {
    await handleLoadClick();
  }
}

/**
 * Reads the currently selected range in Excel and returns the cell values
 * @returns A 2D array of cell values
 */
async function readSelectedRange(): Promise<unknown[][]> {
  return Excel.run(async (context) => {
    // Get the currently selected range
    const range = context.workbook.getSelectedRange();
    
    // Load the values property
    range.load("values");
    
    // Execute the request
    await context.sync();
    
    // Return the 2D array of values
    return range.values;
  });
}

/**
 * Displays the converted datatable in the output area
 */
function displayOutput(datatableText: string): void {
  const outputSection = document.getElementById("output-section");
  const outputElement = document.getElementById("output");
  const copyBtn = document.getElementById("copy-btn") as HTMLButtonElement;
  
  if (outputSection && outputElement) {
    outputElement.textContent = datatableText;
    outputSection.classList.remove("hidden");
  }
  
  if (copyBtn) {
    copyBtn.disabled = false;
  }
}

/**
 * Hides the output section and disables copy button
 */
function hideOutput(): void {
  const outputSection = document.getElementById("output-section");
  const copyBtn = document.getElementById("copy-btn") as HTMLButtonElement;
  const confirmation = document.getElementById("copy-confirmation");
  
  if (outputSection) {
    outputSection.classList.add("hidden");
  }
  
  if (copyBtn) {
    copyBtn.disabled = true;
  }
  
  if (confirmation) {
    confirmation.classList.add("hidden");
  }
  
  lastConvertedResult = "";
}

/**
 * Updates the status message in the task pane
 */
function updateStatus(message: string): void {
  const statusElement = document.getElementById("status");
  if (statusElement) {
    statusElement.textContent = message;
  }
}
