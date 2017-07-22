import * as XLSX from 'xlsx';

import {
  Map,
  CellValue,
  Filter,
  CellFunc,
  Obj,
  Automap,
  Options,
  CellInfo,
  WsInfo,
} from '../src/types';

declare class JsonifyExcel {
  constructor(filePath: string);

  xlsx: typeof XLSX;
  book: XLSX.WorkBook;
  sheet: XLSX.WorkSheet;

  toJson(options: Options, map?: Map): Obj[];
}

declare namespace JsonifyExcel {

}

export = JsonifyExcel;
