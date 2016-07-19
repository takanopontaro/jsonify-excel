const XLSX = require('xlsx');

module.exports = class {
  get defaultOptions() {
    return {
      file: null,
      sheet: 0,
      start: 2,
      condition(cell) { return !!cell('A'); },
    };
  }

  constructor(config, map) {
    const { file, sheet, start, condition } = Object.assign({}, this.defaultOptions, config);
    const workbook = XLSX.readFile(file);
    const sheetName = this.is(sheet, 'string') ? sheet : workbook.SheetNames[sheet];
    this.worksheet = workbook.Sheets[sheetName];
    this.currentRow = start;
    this.condition = condition;
    this.map = map;
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

  collectDate(dataSet = []) {
    const cellValue = this.cellValueGetter(this.currentRow);
    switch (this.condition(cellValue, this.currentRow)) {
      case true:
        dataSet.push(this.parse(this.map, cellValue));
        break;
      case null:
        // skip this turn
        break;
      case false:
      default:
        return dataSet;
    }
    this.currentRow++;
    return this.collectDate(dataSet);
  }

  toJSON() {
    return this.collectDate();
  }
};
