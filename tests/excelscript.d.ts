// Sous-ensemble des types Office Scripts utilisés, mêmes signatures que l'API ExcelScript.
declare namespace ExcelScript {
  interface Workbook { getTable(key: string): Table; }
  interface Table {
    getRangeBetweenHeaderAndTotal(): Range;
    getRowCount(): number;
    deleteRowsAt(index: number, count?: number): void;
    addRows(index?: number, values?: (boolean | string | number)[][]): void;
  }
  interface Range { getValues(): (string | number | boolean)[][]; }
}
