"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var _ = require("lodash");
var XLSX = require("xlsx");
var JsonifyExcel = (function () {
    function JsonifyExcel(filePath) {
        this.xlsx = XLSX;
        this.defaultAutomap = {
            headerRowNum: 0,
            scope: function (value, addr, rowNum, colNum) {
                return true;
            },
        };
        this.defaultOptions = {
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
        this.defaultFilters = {
            trim: function (value, info) {
                if (!_.isString(value)) {
                    return value;
                }
                return value.trim();
            },
            undef: function (value, info) {
                if (value === '') {
                    return undefined;
                }
                return value;
            },
        };
        this.book = XLSX.readFile(filePath, {
            cellFormula: false,
            cellHTML: false,
            cellDates: true,
        });
    }
    JsonifyExcel.prototype.toJson = function (options, map) {
        this.options = _.merge(this.defaultOptions, options);
        var sheet = this.options.sheet;
        var sheetName = _.isString(sheet)
            ? sheet
            : this.book.SheetNames[sheet];
        this.sheet = this.book.Sheets[sheetName];
        this.wsInfo = this.getWsInfo();
        this.initAutomap(map);
        this.adjustStartRowNum('startRowNum' in options);
        this.curRowNum = this.options.startRowNum;
        return this.collectData([]);
    };
    JsonifyExcel.prototype.getWsInfo = function () {
        var ref = this.sheet['!ref'];
        if (ref === undefined) {
            throw new Error('invalid xlsx file detected');
        }
        var _a = XLSX.utils.decode_range(ref), _b = _a.s, sc = _b.c, sr = _b.r, _c = _a.e, ec = _c.c, er = _c.r;
        return {
            startRowNum: sr,
            startColNum: sc,
            endRowNum: er,
            endColNum: ec,
        };
    };
    JsonifyExcel.prototype.initAutomap = function (map) {
        var opts = this.options;
        if (opts.automap === false) {
            this.map = map;
        }
        else {
            opts.automap = _.merge({}, this.defaultAutomap, opts.automap);
            this.map = this.getAutomap();
        }
    };
    JsonifyExcel.prototype.adjustStartRowNum = function (definedByUser) {
        var opts = this.options;
        if (opts.automap === false) {
            return;
        }
        var headerRowNum = opts.automap.headerRowNum;
        if (definedByUser) {
            if (headerRowNum >= opts.startRowNum) {
                throw new Error("startRowNum must be bigger than " + headerRowNum);
            }
        }
        else {
            opts.startRowNum = headerRowNum + 1;
        }
    };
    JsonifyExcel.prototype.getValue = function (value, info) {
        switch (true) {
            case _.isString(value): {
                var md = value.match(/^\*([A-Z]+)$/);
                return md ? info.cell(md[1]) : value;
            }
            case _.isFunction(value):
                return value(info);
            default:
                return value;
        }
    };
    JsonifyExcel.prototype.createCellInfo = function (colNum, col, key) {
        return {
            cell: this.createCellFunc(),
            rowNum: this.curRowNum,
            colNum: colNum,
            col: col,
            key: key,
        };
    };
    JsonifyExcel.prototype.createCellFunc = function () {
        var _this = this;
        return function (colOrKey, automapKey) {
            var col = automapKey === true ? _this.getAutomapCol(colOrKey) : colOrKey;
            return _this.getCellValue("" + col + (_this.curRowNum + 1));
        };
    };
    JsonifyExcel.prototype.isValidRow = function (info) {
        var check = this.options.check;
        if (info.rowNum > this.wsInfo.endRowNum) {
            return false;
        }
        else if (_.isString(check)) {
            return info.cell(check) !== undefined;
        }
        else if (_.isFunction(check)) {
            return check(info);
        }
        else {
            throw new Error('check must be string or function');
        }
    };
    JsonifyExcel.prototype.collectData = function (resultSet) {
        var _this = this;
        var info = this.createCellInfo();
        var check = this.options.check;
        switch (this.isValidRow(info)) {
            case true: {
                var result = _.reduce(this.map, function (res, value, key) {
                    res[key] = _this.getValue(value, info);
                    return res;
                }, {});
                if (this.options.compact) {
                    result = _.omitBy(result, function (v) { return v === undefined; });
                }
                resultSet.push(result);
                break;
            }
            case null:// skip this turn
                break;
            case false:
            default:
                return resultSet;
        }
        this.curRowNum++;
        return this.collectData(resultSet);
    };
    JsonifyExcel.prototype.filterOut = function (value, info) {
        var _this = this;
        var filters = this.options.filters;
        return filters.reduce(function (res, filter) {
            if (_.isString(filter)) {
                return _this.defaultFilters[filter](res, info);
            }
            else if (_.isFunction(filter)) {
                return filter(res, info);
            }
            else {
                throw new Error('filter must be string or function');
            }
        }, value);
    };
    JsonifyExcel.prototype.extractCellValue = function (addr) {
        var c = this.sheet[addr];
        if (c === undefined) {
            return c;
        }
        switch (c.t) {
            case 's':
                return c.w === undefined ? c.v : c.w;
            case 'n':
                if (this.options.number) {
                    return c.v;
                }
                else {
                    return c.w === undefined ? c.v : c.w;
                }
            case 'b':
                return c.v;
            case 'd':
                if (this.options.date) {
                    return new Date(c.v);
                }
                else {
                    return c.w === undefined ? c.v : c.w;
                }
            case 'e':
                if (this.options.error) {
                    return new Error(c.w || '');
                }
                else {
                    return c.w === undefined ? c.v : c.w;
                }
            default:
                return undefined;
        }
    };
    JsonifyExcel.prototype.getCellValue = function (addrOrRowNum, colNum) {
        var addr = _.isString(addrOrRowNum)
            ? addrOrRowNum
            : this.getCellAddr(addrOrRowNum, colNum);
        var value = this.extractCellValue(addr);
        var c = XLSX.utils.decode_range(addr).s.c;
        var col = XLSX.utils.encode_col(c);
        var info = this.createCellInfo(c, col, this.getAutomapKey(col));
        return this.filterOut(value, info);
    };
    JsonifyExcel.prototype.getCellAddr = function (rowNum, colNum) {
        var _a = XLSX.utils, encode_col = _a.encode_col, encode_row = _a.encode_row;
        return "" + encode_col(colNum) + encode_row(rowNum);
    };
    JsonifyExcel.prototype.getAutomapKey = function (col) {
        if (!this.options.automap) {
            return undefined;
        }
        return _.findKey(this.map, function (v) { return v === "*" + col; });
    };
    JsonifyExcel.prototype.getAutomapCol = function (key) {
        if (!this.options.automap) {
            return undefined;
        }
        return this.map[key].replace(/^\*/, '');
    };
    JsonifyExcel.prototype.getAutomap = function () {
        var map = {};
        var _a = this.options.automap, headerRowNum = _a.headerRowNum, scope = _a.scope;
        var _b = this.wsInfo, startColNum = _b.startColNum, endColNum = _b.endColNum;
        for (var colNum = startColNum; colNum <= endColNum; colNum++) {
            var addr = this.getCellAddr(headerRowNum, colNum);
            var _c = XLSX.utils.decode_range(addr).s, c = _c.c, r = _c.r;
            var value = this.extractCellValue(addr);
            if (_.isString(value)) {
                value = value.trim();
            }
            if (scope && !scope(value, addr, r, c)) {
                continue;
            }
            if (value === undefined) {
                return map;
            }
            else if (!_.isString(value)) {
                throw new Error(addr + " must to be string.");
            }
            map[value] = "*" + XLSX.utils.encode_col(colNum);
        }
        return map;
    };
    return JsonifyExcel;
}());
exports.JsonifyExcel = JsonifyExcel;
