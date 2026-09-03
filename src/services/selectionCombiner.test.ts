import { convertToDatatable } from "./datatableConverter";
import { combineSelectedRangeAreas, SelectedRangeArea } from "./selectionCombiner";

function createArea(values: unknown[][], rowIndex: number, columnIndex: number): SelectedRangeArea {
  return {
    values,
    rowIndex,
    columnIndex,
    rowCount: values.length,
    columnCount: values[0]?.length ?? 0,
  };
}

describe("combineSelectedRangeAreas", () => {
  it("rejects an empty selection", () => {
    expect(() => combineSelectedRangeAreas([])).toThrow("No data to convert");
  });

  it("returns a single selected area unchanged", () => {
    const values = [
      ["Name", "Count"],
      ["Alpha", 1],
    ];

    expect(combineSelectedRangeAreas([createArea(values, 0, 0)])).toEqual(values);
  });

  it("combines aligned non-adjacent columns in worksheet order", () => {
    const firstColumn = createArea([["Name"], ["Alpha"], ["Beta"]], 0, 0);
    const sixthColumn = createArea([["Count"], [1], [2]], 0, 5);

    const result = combineSelectedRangeAreas([sixthColumn, firstColumn]);

    expect(result).toEqual([
      ["Name", "Count"],
      ["Alpha", 1],
      ["Beta", 2],
    ]);
  });

  it("preserves column order inside multi-column areas", () => {
    const firstBand = createArea(
      [
        ["Name", "City"],
        ["Alpha", "Seattle"],
      ],
      2,
      0,
    );
    const secondBand = createArea(
      [
        ["Count", "Enabled"],
        [1, true],
      ],
      2,
      4,
    );

    expect(combineSelectedRangeAreas([secondBand, firstBand])).toEqual([
      ["Name", "City", "Count", "Enabled"],
      ["Alpha", "Seattle", 1, true],
    ]);
  });

  it("rejects areas with different starting rows", () => {
    const firstArea = createArea([["Name"], ["Alpha"]], 0, 0);
    const secondArea = createArea([["Count"], [1]], 1, 5);

    expect(() => combineSelectedRangeAreas([firstArea, secondArea])).toThrow(
      "must cover the same rows",
    );
  });

  it("rejects areas with different row counts", () => {
    const firstArea = createArea([["Name"], ["Alpha"]], 0, 0);
    const secondArea = createArea([["Count"], [1], [2]], 0, 5);

    expect(() => combineSelectedRangeAreas([firstArea, secondArea])).toThrow(
      "must cover the same rows",
    );
  });

  it("rejects area values that do not match the reported dimensions", () => {
    const area = createArea([["Name"], ["Alpha"]], 0, 0);
    area.columnCount = 2;

    expect(() => combineSelectedRangeAreas([area])).toThrow(
      "does not match its Excel range dimensions",
    );
  });

  it("rejects overlapping areas", () => {
    const firstArea = createArea(
      [
        ["Name", "City"],
        ["Alpha", "Seattle"],
      ],
      0,
      0,
    );
    const overlappingArea = createArea([["City"], ["Seattle"]], 0, 1);

    expect(() => combineSelectedRangeAreas([firstArea, overlappingArea])).toThrow(
      "must not overlap",
    );
  });

  it("produces data accepted by the datatable converter", () => {
    const names = createArea([["Name"], ["Alpha"], ["Beta"]], 0, 0);
    const counts = createArea([["Count"], [1], [2]], 0, 5);

    const combined = combineSelectedRangeAreas([names, counts]);
    const result = convertToDatatable(combined);

    expect(result).toContain("datatable(Name:string, Count:long)");
    expect(result).toContain('"Alpha", 1,');
    expect(result).toContain('"Beta", 2,');
  });
});
