const fs = require('fs');
const path = require('path').resolve(process.argv[2]) + '/';
const numeric = {fromF64: x => x, fromI32: x => x, fromU32: x => x >>> 0, toF64: x => { if(typeof x !== 'number') throw Error('not number'); return x; }, toI32: x => {if (!Number.isInteger(x) || x < -2147483648 || x > 2147483647) throw Error('not i32'); return x;}, toU32: x => {if (!Number.isInteger(x) || x < 0 || x > 4294967295) throw Error('not u32'); return x;}, test: x => typeof x === 'number', testI32: x => Number.isInteger(x) && x >= -2147483648 && x <= 2147483647, testU32: x => Number.isInteger(x) && x >= 0 && x <= 4294967295};
const supported = {
 'wasm:js-number': numeric,
 'wasm:js-string': {test: x => typeof x === 'string', cast: x => {if(typeof x !== 'string') throw Error('not string');return x;}, fromCharCode: String.fromCharCode, concat: (a,b) => a+b, equals: (a,b) => a===b, compare: (a,b) => a < b ? -1 : a > b ? 1 : 0, length: x => x.length, charCodeAt: (s,i) => {if ((i >>> 0) >= s.length) throw new WebAssembly.RuntimeError('string index out of bounds'); return s.charCodeAt(i >>> 0);}},
 'wasm:js-boolean': {test: x => typeof x === 'boolean', cast: x => {if(typeof x !== 'boolean') throw Error('not bool');return x ? 1 : 0;}},
 'wasm:js-bigint': {test: x => typeof x === 'bigint'},
 'wasm:js-undefined': {test: x => x === undefined},
 'wasm:js-symbol': {test: x => typeof x === 'symbol', equals: (a,b) => a===b},
 'ecma:bigint': {asUintN: BigInt.asUintN, toString: x => x.toString(), BigInt},
 'ecma:number': {Number, parseFloat, isNaN: Number.isNaN},
 'ecma:boolean': {toBoolean: Boolean},
 'ecma:string': {String, fromCharCode: String.fromCharCode, charCodeAt: (s,i) => s.charCodeAt(i)},
 'ecma:map': {new: () => new Map(), set: (m,k,v) => m.set(k,v)},
 'ecma:value': {abstractEq: (a,b) => a == b},
 'ecma:object': {get: (o,k) => o?.[k]},
};
let checks = 0;
for (const format of ['C','n','v','N','V']) {
 const module = new WebAssembly.Module(fs.readFileSync(path + `php-format-${format.charCodeAt(0).toString(16)}.wasm`));
 const imports = {};
 for (const d of WebAssembly.Module.imports(module)) {
  imports[d.module] ??= {};
  imports[d.module][d.name] = d.kind === 'global' ? (d.module === 'wasm:string-constants' ? d.name : d.name === 'true' ? true : d.name === 'false' ? false : undefined) : supported[d.module]?.[d.name] ?? (() => {throw Error(`Unexpected ${d.module}.${d.name}`);});
 }
 const run = new WebAssembly.Instance(module, imports).exports._start;
 for (const [input, bits] of [[-1,4294967295],[-1.5,4294967295],[-257,4294967039],[4294967297,1],[9223372036854775807n,4294967295],[-9223372036854775807n,1],['12junk',12],['-1',4294967295],[true,1],[false,0],[null,0],[NaN,0],[Infinity,0],[-Infinity,0]]) {
  const width = format === 'C' ? 1 : format === 'n' || format === 'v' ? 2 : 4;
  const little = format === 'v' || format === 'V';
  const expected = Array.from({length:width},(_,i) => (bits >>> (8*(little ? i : width-i-1))) & 255);
  const actual = Array.from(run(input), c => c.charCodeAt(0));
  if (JSON.stringify(actual) !== JSON.stringify(expected)) throw Error(`${format}/${input}: ${actual} != ${expected}`);
  checks++;
 }
}
console.log(`${checks} PHP native Wasm packing checks passed`);

let unpackChecks = 0;
for (const format of ['N','V','n','v','Cfirst/Csecond','C*']) {
 const name = Buffer.from(format).toString('hex');
 const module = new WebAssembly.Module(fs.readFileSync(path + `php-unpack-${name}.wasm`));
 const imports = {};
 for (const d of WebAssembly.Module.imports(module)) {
  imports[d.module] ??= {};
  imports[d.module][d.name] = d.kind === 'global' ? (d.module === 'wasm:string-constants' ? d.name : d.name === 'true' ? true : d.name === 'false' ? false : undefined) : supported[d.module]?.[d.name] ?? (() => {throw Error(`Unexpected ${d.module}.${d.name}`);});
 }
 const run = new WebAssembly.Instance(module, imports).exports._start;
 for (const bytes of [[0,0,0,0], [255,128,0,1], [18,52,86,120], [255,255,255,255]]) {
  const actual = Array.from(run(String.fromCharCode(...bytes)));
  let expected;
  if (format === 'C*') expected = bytes.map((b,i) => [i+1,b]);
  else if (format === 'Cfirst/Csecond') expected = [['first',bytes[0]],['second',bytes[1]]];
  else {
   const width = format === 'n' || format === 'v' ? 2 : 4;
   const little = format === 'v' || format === 'V';
   const value = bytes.slice(0,width).reduce((sum,b,i) => sum + b * 2 ** (8*(little ? i : width-i-1)), 0);
   expected = [[1,value]];
  }
  if (JSON.stringify(actual) !== JSON.stringify(expected)) throw Error(`${format}/${bytes}: ${JSON.stringify(actual)} != ${JSON.stringify(expected)}`);
  unpackChecks++;
 }
 if (format === 'C*') {
  if (run('').size !== 0) throw Error('C* empty input');
  unpackChecks++;
 }
}
console.log(`${unpackChecks} PHP native Wasm unpacking checks passed`);
