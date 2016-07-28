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
  automap: false,
  sheet: 0,
  start: 2,
  condition: function (cell) {
    return !!cell('A');
  },
};

const map = [{
  name: '*A',
  retired: '*B',
  born: '*C',
  age: '*D',
  error: '*E',
}];

const json = new Je('test.xlsx').toJSON(config, map);

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

### constructor(filePath)

`filePath` is path to your excel file.

### toJSON(config, map)

Return an array of object or an object based on `config` and `map`.

For details of `config` and `map`, see below.

### config

A plain object has a structure below.

|key|type|default|description|
|---|---|---|---|
|automap|boolean|false|Generate map based on header cells automatically. Cell texts become keys of JSON.|
|sheet|string/number|0|Target sheet name or zero-based index.|
|start|number|2|One-based start row number. If `automap` is `true`, this parameter means header row.|
|condition|function|function (cell, row) { return !!cell('A'); }|Conditional function called just before starting to parse current row. It has 2 arguments. `cell` is function to get a cell value passed column as its arguments. `row` is current row number. It needs to return true (proceed) or false (exit) or null (skip current row).|

### map

A plain object has a structure you want as JSON. Uppercase alphabets start with `*` are replaced with cell data of that column.

If it's wrapped by an array, one row is parsed as one item of returned array. Or all collected data are merged to one object and it's returned.

```js
[{
  name: '*A',
  retired: 'B',
  born: '*c',
  age: '*D',
}]
```

becomes

```js
[ { name: 'Katsuhiro Otomo',
    retired: 'B', // <-- not replaced
    born: '*c', // <-- not replaced
    age: '62' },
  { name: 'Hayao Miyazaki', ... },
  { name: 'Hideaki Anno', ... } ]
```

```js
{
  name: '*A',
}
```

becomes below because the map is object not array and `name` property is overwritten with followed rows.

```js
{ name: 'Hideaki Anno' }
```

You can get cell data as a key of JSON and also use a callback function same as one described in config section above.

```js
{
  '*A': function (cell, row) {
    return cell('A');
  },
}
```

becomes

```js
{
  'Katsuhiro Otomo': 'Katsuhiro Otomo',
  'Hayao Miyazaki': 'Hayao Miyazaki',
  'Hideaki Anno': 'Hideaki Anno'
}
```

#### automap

Mostly, header texts are used as keys of JSON. So you can set `automap` to `true`.

```js
new Je('test.xlsx').toJSON({ automap: true }); // No need 2nd argument
```

is the same map as below.

```js
[{
  Name: '*A',
  Retired: '*B',
  Born: '*C',
  Age: '*D',
  Error: '*E',
}]
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

```bash
$ npm i
$ npm run build
$ npm test
```
