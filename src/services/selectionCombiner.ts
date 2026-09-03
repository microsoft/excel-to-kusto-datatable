/**
 * A rectangular range that is part of the current Excel selection.
 */
export interface SelectedRangeArea {
  values: unknown[][];
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

/**
 * Combines row-aligned selection areas into one rectangular data matrix.
 *
 * Areas are ordered by worksheet column and concatenated horizontally. All areas must cover the
 * same worksheet rows so values from unrelated rows are never paired silently.
 */
export function combineSelectedRangeAreas(areas: SelectedRangeArea[]): unknown[][] {
  if (areas.length === 0) {
    throw new Error("No data to convert. Please select cells with data.");
  }

  areas.forEach(validateArea);

  const expectedRowIndex = areas[0].rowIndex;
  const expectedRowCount = areas[0].rowCount;

  if (
    areas.some((area) => area.rowIndex !== expectedRowIndex || area.rowCount !== expectedRowCount)
  ) {
    throw new Error(
      "Non-adjacent selections must cover the same rows. Select ranges with matching start and end rows.",
    );
  }

  const orderedAreas = [...areas].sort((left, right) => left.columnIndex - right.columnIndex);

  for (let index = 1; index < orderedAreas.length; index++) {
    const previousArea = orderedAreas[index - 1];
    const currentArea = orderedAreas[index];
    const previousEndColumn = previousArea.columnIndex + previousArea.columnCount;

    if (currentArea.columnIndex < previousEndColumn) {
      throw new Error("Selected ranges must not overlap.");
    }
  }

  return Array.from({ length: expectedRowCount }, (_, rowIndex) =>
    orderedAreas.flatMap((area) => area.values[rowIndex]),
  );
}

function validateArea(area: SelectedRangeArea, index: number): void {
  const hasValidPosition =
    Number.isInteger(area.rowIndex) &&
    area.rowIndex >= 0 &&
    Number.isInteger(area.columnIndex) &&
    area.columnIndex >= 0;
  const hasValidDimensions =
    Number.isInteger(area.rowCount) &&
    area.rowCount > 0 &&
    Number.isInteger(area.columnCount) &&
    area.columnCount > 0;
  const valuesMatchDimensions =
    Array.isArray(area.values) &&
    area.values.length === area.rowCount &&
    area.values.every((row) => Array.isArray(row) && row.length === area.columnCount);

  if (!hasValidPosition || !hasValidDimensions || !valuesMatchDimensions) {
    throw new Error(
      `Selected area ${index + 1} does not match its Excel range dimensions. Reselect the cells and try again.`,
    );
  }
}
