# Jsonify Excel

Jsonify Excel files.

## Installation

```bash
$ npm install jsonify-excel
```

## API

Sreadsheet used by this document is below.

`Name` is String, `Retired` is Boolean, `Born` is Date, `Age` is Number and `Error` is Error.

|Name|Retired|Born|Age|Error|
|---|---|---|---|---|
|Katsuhiro Otomo|FALSE|April 14, 1954|62|#DIV/0!|
|Hayao Miyazaki|TRUE|January 5, 1941|75|#NAME?|
|Hideaki Anno|FALSE|May 22, 1960|56|#REF!|

Basic code:

```js
const Je = require('jsonify-excel');

const config = {
  file: 'test.xlsx',
  sheet: 0,
  start: 2,
  condition: function (cell) {
    return !!cell('A');
  },
};

const map = {
  name: '*A',
  retired: '*B',
  born: '*C',
  age: '*D',
  error: '*E',
};

const json = new Je(config, map).toJSON();

console.log(json);
```

becomes

```js
[ { name: 'Katsuhiro Otomo',
    retired: false,
    born: 'April 14, 1954',
    age: '62',
    error: [Error Object] },
  { name: 'Hayao Miyazaki',
    retired: true,
    born: 'January 5, 1941',
    age: '75',
    error: [Error Object] },
  { name: 'Hideaki Anno',
    retired: false,
    born: 'May 22, 1960',
    age: '56',
    error: [Error Object] } ]
```

### constructor(config, map)

For details of `config` and `map`, see below.

### toJSON()

Return Array of object based on `config` and `map`.


### config

A plain object has a structure below.

|key|type|default|description|
|---|---|---|---|
|file|string|null|Path to a excel file|
|sheet|string/number|0|Target sheet name or zero-based index|
|start|number|2|One-based start row number|
|condition|function|function (cell, row) { return !!cell('A'); }|Conditional function called just before starting to parse current row. It has 2 arguments. `cell` is function to get a cell value passed column as its arguments. `row` is current row number. It needs to return true (proceed) or false (exit) or null (skip current row).|

### map

A plain object has a structure you want as JSON.

Uppercase alphabets start with `*` are replaced with cell data of that column.

```js
{
  name: '*A',
  age: '*B',
  address: 'C',
  job: '*d',
}
```

becomes

```js
{
  name: 'Katsuhiro Otomo',
  age: '62',
  address: 'C', // <-- not replaced
  job: '*d', // <-- not replaced
}
```

You can get cell data as a key of JSON and also use a callback function same as one described in config section above.

```js
{
  '*A': 'name',
  name: function (cell, row) {
    return cell('A');
  },
}
```

becomes

```js
{
  'Katsuhiro Otomo': 'name',
  name: 'Katsuhiro Otomo',
}
```

### data type

Returned cell values have data type based on the rules below.

|Excel|JSON|sample|
|---|---|---|
|string|string|'Katsuhiro Otomo'|
|boolean|boolean|true|
|date|string|'April 14, 1954'|
|number|string|'62'|
|error|new Error(cell value)|new Error('#DIV/0!')|

## test

```shell
$ npm i
$ npm run build
$ npm test
```
