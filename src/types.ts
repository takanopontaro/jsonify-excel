export type Map = Obj;
export type CellValue = number | string | boolean | Date | Error;
export type Filter = (value: any, info: CellInfo) => any;
export type CellFunc = (col: string, automapKey?: boolean) => CellValue;

export interface Obj {
  [key: string]: any;
}

export interface Automap {
  headerRowNum?: number;
  scope?: (
    value: CellValue,
    addr: string,
    rowNum: number,
    colNum: number
  ) => boolean;
}

export interface Options {
  automap?: boolean | Automap;
  sheet?: number | string;
  startRowNum?: number;
  check?: string | ((info: CellInfo) => boolean);
  filters?: Array<string | Filter>;
  compact?: boolean;
  number?: boolean;
  date?: boolean;
  error?: boolean;
}

export interface CellInfo {
  cell: CellFunc;
  rowNum: number;
  colNum?: number;
  col?: string;
  key?: string;
}

export interface WsInfo {
  startRowNum: number;
  startColNum: number;
  endRowNum: number;
  endColNum: number;
}
