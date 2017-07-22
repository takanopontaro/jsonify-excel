import * as _ from 'lodash';
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
} from './types';

export default class JsonifyExcel {
  public xlsx = XLSX;
  public book: XLSX.WorkBook;
  public sheet: XLSX.WorkSheet;

  private options: Options;
  private wsInfo: WsInfo;
  private curRowNum: number;
  private map: Map;

  private defaultAutomap: Automap = {
    headerRowNum: 0,
    scope(value, addr, rowNum, colNum) {
      return true;
    },
  };

  private defaultOptions: Options = {
    automap: false,
    sheet: 0,
    startRowNum: 0,
    check: 'A',
    filters: ['trim', 'undef'],
    compact: true,
    number: true,
    date: true,
    error: true,
  };

  private defaultFilters: { [key: string]: Filter } = {
    trim(value, info) {
      if (!_.isString(value)) {
        return value;
      }
      return value.trim();
    },
    undef(value, info) {
      if (value === '') {
        return undefined;
      }
      return value;
    },
  };

  constructor(filePath: string) {
    this.book = XLSX.readFile(filePath, {
      cellFormula: false,
      cellHTML: false,
      cellDates: true,
    });
  }

  public toJson(options: Options, map?: Map): Obj[] {
    this.options = _.merge({}, this.defaultOptions, options);
    const { sheet } = this.options;
    const sheetName = _.isString(sheet) ? sheet : this.book.SheetNames[sheet];
    this.sheet = this.book.Sheets[sheetName];
    this.wsInfo = this.getWsInfo();
    this.initAutomap(map);
    this.adjustStartRowNum('startRowNum' in options);
    this.curRowNum = this.options.startRowNum;
    return this.collectData([]);
  }

  private getWsInfo(): WsInfo {
    const {
      s: { c: sc, r: sr },
      e: { c: ec, r: er },
    } = XLSX.utils.decode_range(this.sheet['!ref']);
    return {
      startRowNum: sr,
      startColNum: sc,
      endRowNum: er,
      endColNum: ec,
    };
  }

  private initAutomap(map: Map): void {
    const opts = this.options;
    if (opts.automap === false) {
      this.map = map;
    } else {
      opts.automap = _.merge({}, this.defaultAutomap, opts.automap);
      this.map = this.getAutomap();
    }
  }

  private adjustStartRowNum(definedByUser: boolean): void {
    const opts = this.options;
    if (opts.automap === false) {
      return;
    }
    const headerRowNum = (opts.automap as Automap).headerRowNum;
    if (definedByUser) {
      if (headerRowNum >= opts.startRowNum) {
        throw new Error(`startRowNum must be bigger than ${headerRowNum}`);
      }
    } else {
      opts.startRowNum = headerRowNum + 1;
    }
  }

  private getValue(value: any, info: CellInfo): any {
    switch (true) {
      case _.isString(value): {
        const md = value.match(/^\*([A-Z]+)$/);
        return md ? info.cell(md[1]) : value;
      }
      case _.isFunction(value):
        return value(info);
      default:
        return value;
    }
  }

  private createCellInfo(
    colNum?: number,
    col?: string,
    key?: string
  ): CellInfo {
    return {
      cell: this.createCellFunc(),
      rowNum: this.curRowNum,
      colNum,
      col,
      key,
    };
  }

  private createCellFunc(): CellFunc {
    return (colOrKey, automapKey) => {
      const col = automapKey === true ? this.getAutomapCol(colOrKey) : colOrKey;
      return this.getCellValue(`${col}${this.curRowNum + 1}`);
    };
  }

  private isValidRow(info: CellInfo): boolean {
    const { check } = this.options;
    if (info.rowNum > this.wsInfo.endRowNum) {
      return false;
    } else if (_.isString(check)) {
      return info.cell(check) !== undefined;
    } else if (_.isFunction(check)) {
      return check(info);
    } else {
      throw new Error('check must be string or function.');
    }
  }

  private collectData(resultSet: Obj[]): Obj[] {
    const info = this.createCellInfo();
    const { check } = this.options;
    switch (this.isValidRow(info)) {
      case true: {
        let result = _.reduce(
          this.map,
          (res, value, key) => {
            res[key] = this.getValue(value, info);
            return res;
          },
          {}
        );
        if (this.options.compact) {
          result = _.omitBy(result, v => v === undefined);
        }
        resultSet.push(result);
        break;
      }
      case null: // skip this turn
        break;
      case false:
      default:
        return resultSet;
    }
    this.curRowNum++;
    return this.collectData(resultSet);
  }

  private filterOut(value: CellValue, info: CellInfo): any {
    return this.options.filters.reduce((res, filter) => {
      if (_.isString(filter)) {
        return this.defaultFilters[filter](res, info);
      } else if (_.isFunction(filter)) {
        return filter(res, info);
      } else {
        throw new Error('filter must be string or function.');
      }
    }, value);
  }

  private extractCellValue(addr: string): CellValue {
    const c = this.sheet[addr];
    if (c === undefined) {
      return c;
    }
    switch (c.t) {
      case 's':
        return c.w === undefined ? c.v : c.w;
      case 'n':
        if (this.options.number) {
          return c.v;
        } else {
          return c.w === undefined ? c.v : c.w;
        }
      case 'b':
        return c.v;
      case 'd':
        if (this.options.date) {
          return new Date(c.v);
        } else {
          return c.w === undefined ? c.v : c.w;
        }
      case 'e':
        if (this.options.error) {
          return new Error(c.w || '');
        } else {
          return c.w === undefined ? c.v : c.w;
        }
      default:
        return undefined;
    }
  }

  private getCellValue(
    addrOrRowNum: string | number,
    colNum?: number
  ): CellValue {
    const addr = _.isString(addrOrRowNum)
      ? addrOrRowNum
      : this.getCellAddr(addrOrRowNum, colNum);
    const value = this.extractCellValue(addr);
    const { s: { c } } = XLSX.utils.decode_range(addr);
    const col = XLSX.utils.encode_col(c);
    const info = this.createCellInfo(c, col, this.getAutomapKey(col));
    return this.filterOut(value, info);
  }

  private getCellAddr(rowNum: number, colNum: number): string {
    const { encode_col, encode_row } = XLSX.utils;
    return `${encode_col(colNum)}${encode_row(rowNum)}`;
  }

  private getAutomapKey(col: string): string {
    if (!this.options.automap) {
      return undefined;
    }
    return _.findKey(this.map, v => v === `*${col}`);
  }

  private getAutomapCol(key: string): string {
    if (!this.options.automap) {
      return undefined;
    }
    return this.map[key].replace(/^\*/, '');
  }

  private getAutomap(): Map {
    const map: Map = {};
    const { headerRowNum, scope } = this.options.automap as Automap;
    const { startColNum, endColNum } = this.wsInfo;
    for (let colNum = startColNum; colNum <= endColNum; colNum++) {
      const addr = this.getCellAddr(headerRowNum, colNum);
      const { s: { c, r } } = XLSX.utils.decode_range(addr);
      let value = this.extractCellValue(addr);
      if (_.isString(value)) {
        value = value.trim();
      }
      if (!scope(value, addr, r, c)) {
        continue;
      }
      if (value === undefined) {
        return map;
      } else if (!_.isString(value)) {
        throw new Error(`${addr} must to be string.`);
      }
      map[value] = `*${XLSX.utils.encode_col(colNum)}`;
    }
    return map;
  }
}
