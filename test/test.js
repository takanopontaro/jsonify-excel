const _ = require('lodash');
const { assert } = require('chai');
const Je = require('../dist/index');

const je = new Je('test/test.xlsx');

const baseConfig = {
  sheet: 0,
  start: 2,
  condition(cell) { return !!cell('A'); },
};

const baseMap = [{
  name: '*A',
  retired: '*B',
  born: '*C',
  age: '*D',
  error: '*E',
}];


describe('actions', function () {
  it('get cell data based on map', function () {
    const json = je.toJSON(baseConfig, baseMap);
    const { name, retired, born, age, error } = json[1];
    assert.equal(json.length, 3);
    assert.equal(name, 'Hayao Miyazaki');
    assert.equal(retired, true);
    assert.equal(born, 'January 5, 1941');
    assert.equal(age, '75');
    assert.equal(error.message, '#NAME?');
  });

  it('should be type suitable for its data', function () {
    const json = je.toJSON(baseConfig, baseMap);
    const { name, retired, born, age, error } = json[0];
    assert.equal(_.isString(name), true);
    assert.equal(_.isBoolean(retired), true);
    assert.equal(_.isString(born), true);
    assert.equal(_.isString(age), true);
    assert.equal(_.isError(error), true);
  });

  it('should be skipped row 3', function () {
    const config = _.merge({}, baseConfig, {
      condition(cell, row) {
        if (row === 3) return null;
        return !!cell('A');
      }
    });
    const json = je.toJSON(config, baseMap);
    assert.equal(json.length, 2);
    assert.equal(json[0].name, 'Katsuhiro Otomo');
    assert.equal(json[1].name, 'Hideaki Anno');
  });

  it('key of json is dynamic and value is possible to be function', function () {
    const map = {
      '*A': (cell, row) => `${row}: ${cell('C')} (${cell('D')})`,
    };
    const json = je.toJSON(baseConfig, map);
    assert.equal(json['Katsuhiro Otomo'], '2: April 14, 1954 (62)');
  });
});
