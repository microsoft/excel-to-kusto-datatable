/**
 * Unit tests for datatableConverter.ts
 */

import {
  convertToDatatable,
  inferColumns,
  KustoDataType,
  KUSTO_DATA_TYPES,
} from "./datatableConverter";

describe("convertToDatatable", () => {
  describe("Basic conversion", () => {
    it("should convert simple string data with headers", () => {
      const data = [
        ["Name", "City"],
        ["Alice", "Seattle"],
        ["Bob", "Portland"],
      ];

      const result = convertToDatatable(data);

      expect(result).toContain("datatable(Name:string, City:string)");
      expect(result).toContain('"Alice", "Seattle",');
      expect(result).toContain('"Bob", "Portland",');
    });

    it("should convert data without headers when option is false", () => {
      const data = [
        ["Alice", "Seattle"],
        ["Bob", "Portland"],
      ];

      const result = convertToDatatable(data, { firstRowIsHeader: false });

      expect(result).toContain("datatable(Column1:string, Column2:string)");
      expect(result).toContain('"Alice", "Seattle",');
      expect(result).toContain('"Bob", "Portland",');
    });
  });

  describe("Type inference", () => {
    it("should infer long type for integer numbers", () => {
      const data = [["Count"], [1], [42], [100]];

      const result = convertToDatatable(data);

      expect(result).toContain("Count:long");
      expect(result).toContain("1,");
      expect(result).toContain("42,");
      expect(result).toContain("100,");
    });

    it("should infer real type for decimal numbers", () => {
      const data = [["Price"], [10.5], [20.99], [0.5]];

      const result = convertToDatatable(data);

      expect(result).toContain("Price:real");
      expect(result).toContain("10.5,");
      expect(result).toContain("20.99,");
    });

    it("should use real when column has mixed integers and decimals", () => {
      const data = [["Value"], [10], [20.5], [30]];

      const result = convertToDatatable(data);

      expect(result).toContain("Value:real");
    });

    it("should infer datetime type for date values", () => {
      const data = [["Date"], ["2024-01-15"], ["2024-02-20"]];

      const result = convertToDatatable(data);

      expect(result).toContain("Date:datetime");
      expect(result).toContain("datetime(2024-01-15)");
    });

    it("should infer string type when column has mixed types", () => {
      const data = [["Mixed"], [123], ["hello"], [456]];

      const result = convertToDatatable(data);

      expect(result).toContain("Mixed:string");
    });
  });

  describe("Null and empty handling", () => {
    it("should handle null values in string columns", () => {
      const data = [["Name"], ["Alice"], [null], ["Bob"]];

      const result = convertToDatatable(data);

      expect(result).toContain("Name:string");
      expect(result).toContain('"",'); // Empty string for null in string column
    });

    it("should handle null values in numeric columns", () => {
      const data = [["Count"], [10], [null], [20]];

      const result = convertToDatatable(data);

      expect(result).toContain("Count:long");
      expect(result).toContain("long(null),");
    });

    it("should handle empty string values", () => {
      const data = [["Name"], ["Alice"], [""], ["Bob"]];

      const result = convertToDatatable(data);

      expect(result).toContain('"",'); // Empty string preserved
    });

    it("should handle null values in real columns", () => {
      const data = [["Price"], [10.5], [null], [20.5]];

      const result = convertToDatatable(data);

      expect(result).toContain("Price:real");
      expect(result).toContain("real(null),");
    });

    it("should handle null values in datetime columns", () => {
      const data = [["Date"], ["2024-01-15"], [null], ["2024-02-20"]];

      const result = convertToDatatable(data);

      expect(result).toContain("Date:datetime");
      expect(result).toContain("datetime(null),");
    });
  });

  describe("String escaping", () => {
    it("should escape double quotes in string values", () => {
      const data = [["Quote"], ['He said "hello"']];

      const result = convertToDatatable(data);

      expect(result).toContain('"He said \\"hello\\"",');
    });

    it("should escape backslashes in string values", () => {
      const data = [["Path"], ["C:\\Users\\Test"]];

      const result = convertToDatatable(data);

      expect(result).toContain('"C:\\\\Users\\\\Test",');
    });
  });

  describe("Column name sanitization", () => {
    it("should replace spaces in column names with underscores", () => {
      const data = [
        ["First Name", "Last Name"],
        ["Alice", "Smith"],
      ];

      const result = convertToDatatable(data);

      expect(result).toContain("First_Name:string");
      expect(result).toContain("Last_Name:string");
    });

    it("should prefix numeric column names with underscore", () => {
      const data = [["123Column"], ["value"]];

      const result = convertToDatatable(data);

      expect(result).toContain("_123Column:string");
    });

    it("should replace special characters in column names", () => {
      const data = [["Name!@#"], ["value"]];

      const result = convertToDatatable(data);

      expect(result).toContain("Name___:string");
    });
  });

  describe("Edge cases", () => {
    it("should throw error for empty data", () => {
      expect(() => convertToDatatable([])).toThrow("No data to convert");
    });

    it("should throw error when only header row exists", () => {
      const data = [["Header1", "Header2"]];

      expect(() => convertToDatatable(data)).toThrow("No data rows to convert");
    });

    it("should handle single cell", () => {
      const data = [["Header"], ["Value"]];

      const result = convertToDatatable(data);

      expect(result).toContain("datatable(Header:string)");
      expect(result).toContain('"Value",');
    });

    it("should handle numeric strings as numbers", () => {
      const data = [["Number"], ["123"], ["456.78"]];

      const result = convertToDatatable(data);

      // Since there's a decimal, should be real
      expect(result).toContain("Number:real");
    });
  });

  describe("Date formats", () => {
    it("should handle ISO date format", () => {
      const data = [["Date"], ["2024-01-15"]];

      const result = convertToDatatable(data);

      expect(result).toContain("datetime(2024-01-15)");
    });

    it("should handle US date format", () => {
      const data = [["Date"], ["01/15/2024"]];

      const result = convertToDatatable(data);

      expect(result).toContain("datetime(01/15/2024)");
    });

    it("should handle Excel serial dates (numbers) as long by default since they cannot be distinguished from integers", () => {
      const data = [
        ["Date"],
        [45306], // Excel serial date for 2024-01-15 - looks like an integer
      ];

      const result = convertToDatatable(data, { firstRowIsHeader: true });

      // Without explicit dateColumns option, serial dates look like integers
      expect(result).toContain("Date:long");
      expect(result).toContain("45306,");
    });
  });

  describe("columnTypes option", () => {
    it("should convert Excel serial dates to datetime when column type is specified", () => {
      const data = [
        ["StartDate", "Value"],
        [45306, 100], // Excel serial date for 2024-01-15
        [45337, 200], // Excel serial date for 2024-02-15
      ];

      const columnTypes = new Map<number, KustoDataType>([[0, "datetime"]]);
      const result = convertToDatatable(data, { columnTypes });

      expect(result).toContain("StartDate:datetime");
      expect(result).toContain("Value:long");
      expect(result).toContain("datetime(2024-01-15");
      expect(result).toContain("datetime(2024-02-15");
    });

    it("should handle multiple column type overrides", () => {
      const data = [
        ["StartDate", "Value", "EndDate"],
        [45306, 100, 45337],
        [45337, 200, 45368],
      ];

      const columnTypes = new Map<number, KustoDataType>([
        [0, "datetime"],
        [2, "datetime"],
      ]);
      const result = convertToDatatable(data, { columnTypes });

      expect(result).toContain("StartDate:datetime");
      expect(result).toContain("Value:long");
      expect(result).toContain("EndDate:datetime");
    });

    it("should handle null values in datetime columns", () => {
      const data = [
        ["Date", "Value"],
        [45306, 100],
        [null, 200],
        [45337, 300],
      ];

      const columnTypes = new Map<number, KustoDataType>([[0, "datetime"]]);
      const result = convertToDatatable(data, { columnTypes });

      expect(result).toContain("Date:datetime");
      expect(result).toContain("datetime(null),");
    });

    it("should allow overriding inferred types to any Kusto type", () => {
      const data = [
        ["Id", "Active", "Data"],
        ["abc-123", "true", '{"key": "value"}'],
        ["def-456", "false", '{"key": "other"}'],
      ];

      const columnTypes = new Map<number, KustoDataType>([
        [0, "guid"],
        [1, "bool"],
        [2, "dynamic"],
      ]);
      const result = convertToDatatable(data, { columnTypes });

      expect(result).toContain("Id:guid");
      expect(result).toContain("Active:bool");
      expect(result).toContain("Data:dynamic");
      expect(result).toContain("guid(abc-123)");
      expect(result).toContain("true,");
      expect(result).toContain('dynamic({"key": "value"})');
    });

    it("should support decimal type", () => {
      const data = [["Amount"], [123.45], [678.9]];

      const columnTypes = new Map<number, KustoDataType>([[0, "decimal"]]);
      const result = convertToDatatable(data, { columnTypes });

      expect(result).toContain("Amount:decimal");
      expect(result).toContain("decimal(123.45)");
    });

    it("should support timespan type", () => {
      const data = [["Duration"], ["01:30:00"], ["02:45:30"]];

      const columnTypes = new Map<number, KustoDataType>([[0, "timespan"]]);
      const result = convertToDatatable(data, { columnTypes });

      expect(result).toContain("Duration:timespan");
      expect(result).toContain("timespan(01:30:00)");
    });
  });

  describe("inferColumns", () => {
    it("should infer column types from data", () => {
      const data = [
        ["Name", "Age", "Salary"],
        ["Alice", 30, 75000.5],
        ["Bob", 25, 65000.0],
      ];

      const columns = inferColumns(data, true);

      expect(columns).toHaveLength(3);
      expect(columns[0]).toEqual({ name: "Name", type: "string" });
      expect(columns[1]).toEqual({ name: "Age", type: "long" });
      expect(columns[2]).toEqual({ name: "Salary", type: "real" });
    });

    it("should use default headers when firstRowIsHeader is false", () => {
      const data = [
        ["Alice", 30],
        ["Bob", 25],
      ];

      const columns = inferColumns(data, false);

      expect(columns).toHaveLength(2);
      expect(columns[0].name).toBe("Column1");
      expect(columns[1].name).toBe("Column2");
    });

    it("should return empty array for empty data", () => {
      const columns = inferColumns([], true);
      expect(columns).toHaveLength(0);
    });

    it("should handle header-only data", () => {
      const data = [["Col1", "Col2"]];
      const columns = inferColumns(data, true);

      expect(columns).toHaveLength(2);
      expect(columns[0].type).toBe("string"); // Default when no data rows
    });
  });

  describe("KUSTO_DATA_TYPES", () => {
    it("should contain all expected types", () => {
      expect(KUSTO_DATA_TYPES).toContain("string");
      expect(KUSTO_DATA_TYPES).toContain("long");
      expect(KUSTO_DATA_TYPES).toContain("real");
      expect(KUSTO_DATA_TYPES).toContain("datetime");
      expect(KUSTO_DATA_TYPES).toContain("bool");
      expect(KUSTO_DATA_TYPES).toContain("guid");
      expect(KUSTO_DATA_TYPES).toContain("timespan");
      expect(KUSTO_DATA_TYPES).toContain("decimal");
      expect(KUSTO_DATA_TYPES).toContain("dynamic");
    });
  });

  describe("Multi-column mixed types", () => {
    it("should handle multiple columns with different types", () => {
      const data = [
        ["Name", "Age", "Salary", "JoinDate"],
        ["Alice", 30, 75000.5, "2020-01-15"],
        ["Bob", 25, 65000.0, "2021-06-01"],
      ];

      const result = convertToDatatable(data);

      expect(result).toContain("Name:string");
      expect(result).toContain("Age:long");
      expect(result).toContain("Salary:real");
      expect(result).toContain("JoinDate:datetime");
    });
  });

  describe("Output format", () => {
    it("should have correct datatable structure", () => {
      const data = [["Col1"], ["val1"]];

      const result = convertToDatatable(data);
      const lines = result.split("\n");

      expect(lines[0]).toMatch(/^datatable\(.+\)$/);
      expect(lines[1]).toBe("[");
      expect(lines[lines.length - 1]).toBe("]");
    });

    it("should indent data rows", () => {
      const data = [["Col1"], ["val1"]];

      const result = convertToDatatable(data);

      expect(result).toContain('    "val1",');
    });

    it("should add trailing comma to each row", () => {
      const data = [["Col1"], ["val1"], ["val2"]];

      const result = convertToDatatable(data);

      // Each data row should end with a comma
      const lines = result.split("\n");
      const dataLines = lines.filter((l) => l.startsWith("    "));
      dataLines.forEach((line) => {
        expect(line.endsWith(",")).toBe(true);
      });
    });
  });
});
