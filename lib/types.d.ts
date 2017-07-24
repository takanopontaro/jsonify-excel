export declare type Map = any;
export declare type CellValue = number | string | boolean | Date | Error;
export declare type Filter = (value: any, info: CellInfo) => any;
export declare type CellFunc = (col: string, automapKey?: boolean) => CellValue;
export interface Automap {
    headerRowNum?: number;
    scope?: (value: CellValue | undefined, addr: string, rowNum: number, colNum: number) => boolean;
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
