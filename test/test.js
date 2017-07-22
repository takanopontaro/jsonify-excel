import test from 'ava';
import Je from '..';

const je = new Je('test/test.xlsx');

test('user map', t => {
  const json = je.toJson(
    { startRowNum: 1 },
    {
      Name: '*A',
      Retired: '*B',
      Born: '*C',
      Age: '*D',
      Error: '*E',
      Sex: '*G',
    },
  );
  t.is(json[1].Age, 75);
});

test('automap', t => {
  const json = je.toJson({ automap: true });
  t.is(json.length, 3);
  t.is(Object.keys(json[0]).length, 5);
});

test('scope', t => {
  const json = je.toJson({
    automap: {
      scope(value, addr, rowNum, colNum) {
        return value !== undefined;
      },
    },
  });
  t.is(Object.keys(json[0]).length, 6);
});

test('data type', t => {
  const json = je.toJson({ automap: true });
  t.is(json[0].Name.constructor, String);
  t.is(json[0].Retired.constructor, Boolean);
  t.is(json[0].Born.constructor, Date);
  t.is(json[0].Age.constructor, Number);
  t.is(json[0].Error.constructor, Error);
});

test('compact', t => {
  let json = je.toJson({ automap: true });
  t.is('Error' in json[1], false);
  json = je.toJson({ automap: true, compact: false });
  t.is('Error' in json[1], true);
});

test('number', t => {
  let json = je.toJson({ automap: true });
  t.is(json[2].Age, 56.5);
  json = je.toJson({ automap: true, number: false });
  t.is(json[2].Age, '56.5');
});

test('filter', t => {
  const json = je.toJson({
    automap: true,
    filters: [
      'trim',
      'undef',
      (value, info) => {
        if (
          info.rowNum === 2 &&
          info.col === 'C' &&
          info.colNum === 2 &&
          info.key === 'Born' &&
          info.cell('Name', true) === 'Hayao Miyazaki'
        ) {
          return value.getFullYear();
        }
        return value;
      },
    ],
  });
  t.is(json[1].Born, 1941);
});
