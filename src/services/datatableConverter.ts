/**
 * datatableConverter.ts
 *
 * Pure conversion logic for transforming Excel cell data into Kusto datatable syntax.
 * No Office.js dependencies - this module can be tested independently.
 */

/**
 * Kusto data types supported by the converter
 */
export type KustoDataType =
  | "string"
  | "long"
  | "real"
  | "datetime"
  | "bool"
  | "guid"
  | "timespan"
  | "decimal"
  | "dynamic";

/**
 * All available Kusto data types for UI selection
 */
export const KUSTO_DATA_TYPES: KustoDataType[] = [
  "string",
  "long",
  "real",
  "datetime",
  "bool",
  "guid",
  "timespan",
  "decimal",
  "dynamic",
];

/**
 * Represents a column with its inferred type
 */
export interface ColumnInfo {
  name: string;
  type: KustoDataType;
}

/**
 * Options for datatable conversion
 */
export interface ConversionOptions {
  /** If true, first row is treated as column headers. Default: true */
  firstRowIsHeader: boolean;
  /**
   * Explicit type overrides for columns, keyed by column index.
   * If provided, these types will be used instead of inferred types.
   */
  columnTypes?: Map<number, KustoDataType>;
}

/**
 * Default conversion options
 */
const DEFAULT_OPTIONS: ConversionOptions = {
  firstRowIsHeader: true,
};

/**
 * Infers column information (names and types) from Excel data without converting.
 * Use this to display column mappings to the user before conversion.
 *
 * @param data - 2D array of cell values from Excel
 * @param firstRowIsHeader - Whether the first row contains headers
 * @returns Array of ColumnInfo with inferred types
 */
export function inferColumns(data: unknown[][], firstRowIsHeader: boolean = true): ColumnInfo[] {
  if (!data || data.length === 0) {
    return [];
  }

  const headers = firstRowIsHeader ? data[0] : generateDefaultHeaders(data[0].length);
  const dataRows = firstRowIsHeader ? data.slice(1) : data;

  if (dataRows.length === 0) {
    // No data rows, just return headers with string type
    return headers.map((header, index) => ({
      name: sanitizeColumnName(String(header ?? `Column${index + 1}`)),
      type: "string" as KustoDataType,
    }));
  }

  return inferColumnTypes(headers, dataRows);
}

/**
 * Converts a 2D array of Excel cell values to Kusto datatable syntax
 *
 * @param data - 2D array of cell values from Excel
 * @param options - Conversion options
 * @returns Kusto datatable syntax as a string
 */
export function convertToDatatable(
  data: unknown[][],
  options: Partial<ConversionOptions> = {},
): string {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  // Validate input
  if (!data || data.length === 0) {
    throw new Error("No data to convert. Please select cells with data.");
  }

  // Separate headers and data rows
  const headers = opts.firstRowIsHeader ? data[0] : generateDefaultHeaders(data[0].length);
  const dataRows = opts.firstRowIsHeader ? data.slice(1) : data;

  if (dataRows.length === 0) {
    throw new Error("No data rows to convert. If first row is header, select at least 2 rows.");
  }

  // Infer column types from all data rows, with optional overrides
  const columns = inferColumnTypes(headers, dataRows, opts.columnTypes);

  // Build the datatable string
  return buildDatatableString(columns, dataRows);
}

/**
 * Generates default column headers (Column1, Column2, etc.)
 */
function generateDefaultHeaders(count: number): string[] {
  return Array.from({ length: count }, (_, i) => `Column${i + 1}`);
}

/**
 * Infers the Kusto data type for each column based on all values in that column
 * @param headers - Column headers
 * @param dataRows - Data rows
 * @param columnTypes - Optional map of column index to explicit type override
 */
function inferColumnTypes(
  headers: unknown[],
  dataRows: unknown[][],
  columnTypes?: Map<number, KustoDataType>,
): ColumnInfo[] {
  return headers.map((header, colIndex) => {
    const columnValues = dataRows.map((row) => row[colIndex]);

    // Use explicit type if provided, otherwise infer
    const type = columnTypes?.get(colIndex) ?? inferTypeFromValues(columnValues);

    return {
      name: sanitizeColumnName(String(header ?? `Column${colIndex + 1}`)),
      type,
    };
  });
}

/**
 * Sanitizes a column name for Kusto (removes invalid characters, ensures valid identifier)
 */
function sanitizeColumnName(name: string): string {
  // Trim whitespace
  let sanitized = name.trim();

  // If empty, use a default
  if (!sanitized) {
    return "Column";
  }

  // Replace spaces and invalid characters with underscores
  sanitized = sanitized.replace(/[^a-zA-Z0-9_]/g, "_");

  // Ensure it doesn't start with a number
  if (/^[0-9]/.test(sanitized)) {
    sanitized = "_" + sanitized;
  }

  return sanitized;
}

/**
 * Infers the most appropriate Kusto type from a set of column values.
 * Uses a priority system: datetime > real > long > string
 */
function inferTypeFromValues(values: unknown[]): KustoDataType {
  // Track what types we've seen (excluding nulls/empty)
  let hasLong = false;
  let hasReal = false;
  let hasDatetime = false;
  let hasString = false;

  for (const value of values) {
    // Skip null/empty values - they're compatible with any type
    if (isNullOrEmpty(value)) {
      continue;
    }

    const valueType = inferSingleValueType(value);
    switch (valueType) {
      case "long":
        hasLong = true;
        break;
      case "real":
        hasReal = true;
        break;
      case "datetime":
        hasDatetime = true;
        break;
      case "string":
        hasString = true;
        break;
    }
  }

  // If any value is a non-date string, the whole column must be string
  if (hasString) {
    return "string";
  }

  // If we have dates, use datetime (dates don't mix well with numbers)
  if (hasDatetime) {
    return "datetime";
  }

  // If we have any decimal numbers, use real
  if (hasReal) {
    return "real";
  }

  // If we only have integers, use long
  if (hasLong) {
    return "long";
  }

  // All values were null/empty - default to string
  return "string";
}

/**
 * Infers the type of a single value
 */
function inferSingleValueType(value: unknown): KustoDataType {
  if (isNullOrEmpty(value)) {
    return "string"; // Null is compatible with any type, default to string
  }

  // Check if it's a number
  if (typeof value === "number") {
    if (!isFinite(value)) {
      return "string"; // NaN or Infinity → treat as string
    }
    return Number.isInteger(value) ? "long" : "real";
  }

  // Check if it's a string that might be a date or number
  if (typeof value === "string") {
    const trimmed = value.trim();

    // Try parsing as a date first
    if (isDateString(trimmed)) {
      return "datetime";
    }

    // Try parsing as a number
    const num = Number(trimmed);
    if (!isNaN(num) && trimmed !== "") {
      return Number.isInteger(num) ? "long" : "real";
    }

    return "string";
  }

  // Check for Date objects (Excel might return these)
  if (value instanceof Date) {
    return "datetime";
  }

  // Everything else is a string
  return "string";
}

/**
 * Checks if a string appears to be a date.
 * We're intentionally permissive here since Kusto's datetime() can parse many formats.
 */
function isDateString(value: string): boolean {
  // Common date patterns
  const datePatterns = [
    // ISO format: 2024-01-15, 2024-01-15T10:30:00
    /^\d{4}-\d{2}-\d{2}(T[\d:]+)?/,
    // US format: 01/15/2024, 1/15/2024
    /^\d{1,2}\/\d{1,2}\/\d{2,4}$/,
    // European format: 15-01-2024, 15.01.2024
    /^\d{1,2}[-.](\d{1,2})[-.](\d{2,4})$/,
  ];

  for (const pattern of datePatterns) {
    if (pattern.test(value)) {
      // Verify it actually parses as a valid date
      const parsed = Date.parse(value);
      if (!isNaN(parsed)) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Checks if a value is null, undefined, or empty string
 */
function isNullOrEmpty(value: unknown): boolean {
  return value === null || value === undefined || value === "";
}

/**
 * Builds the complete datatable string from columns and data
 */
function buildDatatableString(columns: ColumnInfo[], dataRows: unknown[][]): string {
  const lines: string[] = [];

  // Build schema declaration
  const schemaItems = columns.map((col) => `${col.name}:${col.type}`);
  lines.push(`datatable(${schemaItems.join(", ")})`);
  lines.push("[");

  // Build data rows
  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
    const row = dataRows[rowIndex];
    const formattedValues = columns.map((col, colIndex) => formatValue(row[colIndex], col.type));

    const rowStr = "    " + formattedValues.join(", ");
    // Add comma after each row (including the last one, per Kusto syntax)
    lines.push(rowStr + ",");
  }

  lines.push("]");

  return lines.join("\n");
}

/**
 * Formats a single value according to its Kusto type
 */
function formatValue(value: unknown, type: KustoDataType): string {
  // Handle null/empty values with typed nulls
  if (isNullOrEmpty(value)) {
    return formatNull(type);
  }

  switch (type) {
    case "string":
      return formatString(value);
    case "long":
      return formatLong(value);
    case "real":
      return formatReal(value);
    case "datetime":
      return formatDatetime(value);
    case "bool":
      return formatBool(value);
    case "guid":
      return formatGuid(value);
    case "timespan":
      return formatTimespan(value);
    case "decimal":
      return formatDecimal(value);
    case "dynamic":
      return formatDynamic(value);
    default:
      return formatString(value);
  }
}

/**
 * Returns the typed null representation for a Kusto type
 */
function formatNull(type: KustoDataType): string {
  switch (type) {
    case "string":
      return '""';
    case "long":
      return "long(null)";
    case "real":
      return "real(null)";
    case "datetime":
      return "datetime(null)";
    case "bool":
      return "bool(null)";
    case "guid":
      return "guid(null)";
    case "timespan":
      return "timespan(null)";
    case "decimal":
      return "decimal(null)";
    case "dynamic":
      return "dynamic(null)";
    default:
      return '""';
  }
}

/**
 * Formats a value as a Kusto string literal
 */
function formatString(value: unknown): string {
  const str = String(value);
  // Escape backslashes first, then double quotes
  const escaped = str.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  return `"${escaped}"`;
}

/**
 * Formats a value as a Kusto long (integer)
 */
function formatLong(value: unknown): string {
  if (typeof value === "number") {
    return Math.round(value).toString();
  }
  const num = Number(value);
  if (isNaN(num)) {
    return "long(null)";
  }
  return Math.round(num).toString();
}

/**
 * Formats a value as a Kusto real (floating point)
 */
function formatReal(value: unknown): string {
  if (typeof value === "number") {
    return value.toString();
  }
  const num = Number(value);
  if (isNaN(num)) {
    return "real(null)";
  }
  return num.toString();
}

/**
 * Formats a value as a Kusto datetime
 * For ambiguous dates, we pass through to datetime() operator as-is
 */
function formatDatetime(value: unknown): string {
  // If it's a Date object, format as ISO
  if (value instanceof Date) {
    return `datetime(${value.toISOString()})`;
  }

  // If it's a number (Excel serial date), convert it
  if (typeof value === "number") {
    const date = excelSerialToDate(value);
    return `datetime(${date.toISOString()})`;
  }

  // If it's a string, pass through to datetime() - let Kusto handle parsing
  // This handles ambiguous formats like "01/07/2026"
  const str = String(value).trim();
  return `datetime(${str})`;
}

/**
 * Converts an Excel serial date number to a JavaScript Date
 * Excel's epoch is January 1, 1900 (with the infamous Lotus 123 leap year bug)
 */
function excelSerialToDate(serial: number): Date {
  // Excel epoch: December 30, 1899 (accounting for the leap year bug)
  const excelEpoch = new Date(1899, 11, 30);
  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  return new Date(excelEpoch.getTime() + serial * millisecondsPerDay);
}

/**
 * Formats a value as a Kusto bool
 */
function formatBool(value: unknown): string {
  if (typeof value === "boolean") {
    return value ? "true" : "false";
  }
  const str = String(value).toLowerCase().trim();
  if (str === "true" || str === "1" || str === "yes") {
    return "true";
  }
  if (str === "false" || str === "0" || str === "no") {
    return "false";
  }
  return "bool(null)";
}

/**
 * Formats a value as a Kusto guid
 */
function formatGuid(value: unknown): string {
  const str = String(value).trim();
  return `guid(${str})`;
}

/**
 * Formats a value as a Kusto timespan
 */
function formatTimespan(value: unknown): string {
  const str = String(value).trim();
  return `timespan(${str})`;
}

/**
 * Formats a value as a Kusto decimal
 */
function formatDecimal(value: unknown): string {
  if (typeof value === "number") {
    return `decimal(${value})`;
  }
  const num = Number(value);
  if (isNaN(num)) {
    return "decimal(null)";
  }
  return `decimal(${num})`;
}

/**
 * Formats a value as a Kusto dynamic (JSON)
 */
function formatDynamic(value: unknown): string {
  if (typeof value === "object" && value !== null) {
    try {
      return `dynamic(${JSON.stringify(value)})`;
    } catch {
      return `dynamic(${formatString(value)})`;
    }
  }
  // For strings, check if it looks like JSON
  const str = String(value).trim();
  if ((str.startsWith("{") && str.endsWith("}")) || (str.startsWith("[") && str.endsWith("]"))) {
    return `dynamic(${str})`;
  }
  // Wrap non-JSON values as a dynamic string
  return `dynamic(${formatString(value)})`;
}
