const XLSX = require('xlsx');

module.exports = class {
  get defaultOptions() {
    return {
      automap: false,
      sheet: 0,
      start: 2,
      condition(cell) { return !!cell('A'); },
      filter: null,
    };
  }

  constructor(filePath) {
    this.workbook = XLSX.readFile(filePath);
  }

  is(value, type) {
    return Object.prototype.toString.call(value).match(/ (.+)]/)[1].toLowerCase() === type;
  }

  parse(value, cellValue) {
    switch (true) {
      case this.is(value, 'string'): {
        const md = value.match(/^\*([A-Z]+)$/);
        return md ? cellValue(md[1]) : value;
      }
      case this.is(value, 'object'): {
        const data = {};
        Object.keys(value).forEach(key => {
          data[this.parse(key, cellValue)] = this.parse(value[key], cellValue);
        });
        return data;
      }
      case this.is(value, 'array'):
        return value.map(v => this.parse(v, cellValue));
      case this.is(value, 'function'):
        return value(cellValue, this.currentRow);
      default:
        return value;
    }
  }

  cellValueGetter(row) {
    return col => {
      const cell = this.worksheet[col + row];
      if (cell === undefined) return cell;
      switch (cell.t) {
        case 'b':
          return cell.v;
        case 'e':
          return new Error(cell.w);
        default:
          return cell.w;
      }
    };
  }

  collectData(dataSet, map) {
    const cellValue = this.cellValueGetter(this.currentRow);
    switch (this.condition(cellValue, this.currentRow)) {
      case true: {
        const data = this.parse(map, cellValue);
        if (dataSet.push) dataSet.push(data);
        else Object.assign(dataSet, data);
        break;
      }
      case null:
        // skip this turn
        break;
      case false:
      default:
        return dataSet;
    }
    this.currentRow++;
    return this.collectData(dataSet, map);
  }

  getHeaderMap() {
    const cellValue = this.cellValueGetter(this.currentRow++);
    const map = {};
    for (let i = 0; ; i++) {
      const col = XLSX.utils.encode_col(i);
      const value = cellValue(col);
      if (!value) return [map];
      if (this.filter && !this.filter(col)) continue;
      map[value] = `*${col}`;
    }
  }

  toJSON(config, map) {
    const { automap, sheet, start, condition, filter } =
      Object.assign({}, this.defaultOptions, config);
    const sheetName = this.is(sheet, 'string') ? sheet : this.workbook.SheetNames[sheet];
    const start2 = (automap && config.start === undefined) ? 1 : start;
    this.worksheet = this.workbook.Sheets[sheetName];
    this.currentRow = start2;
    this.condition = condition;
    this.filter = filter;
    this.map = automap ? this.getHeaderMap() : map;
    return this.is(this.map, 'array') ?
      this.collectData([], this.map[0]) : this.collectData({}, this.map);
  }
};
