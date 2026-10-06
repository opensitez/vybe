const fs = require('fs');
// Run export_packing_parameter_fixtures first, then pass its output directory.
const path = require('path').resolve(process.argv[2]) + '/';
let checks = 0;
for (const width of [1, 2, 4]) for (const endian of ['Big', 'Little']) {
  const module = new WebAssembly.Module(fs.readFileSync(path + `width-${width}-${endian}.wasm`));
  const imports = {};
  for (const d of WebAssembly.Module.imports(module)) {
    imports[d.module] ??= {};
    if (d.kind === 'global') imports[d.module][d.name] = d.name === 'true' ? true : d.name === 'false' ? false : undefined;
    else imports[d.module][d.name] = (...args) => {
      if (d.module === 'wasm:js-number') {
        const value = args[0];
        if (d.name === 'fromU32') return value >>> 0;
        if (['fromF64', 'fromI32', 'toF64'].includes(d.name)) return value;
        if (d.name === 'toI32') { if (!Number.isInteger(value) || value < -2147483648 || value > 2147483647) throw Error('checked i32'); return value; }
        if (d.name === 'toU32') { if (!Number.isInteger(value) || value < 0 || value > 4294967295) throw Error('checked u32'); return value; }
      }
      if (d.name === 'fromCharCode') return String.fromCharCode(args[0]);
      if (d.name === 'concat') return args[0] + args[1];
      throw Error(`Unexpected import ${d.module}.${d.name}`);
    };
  }
  const run = new WebAssembly.Instance(module, imports).exports._start;
  const values = width === 1 ? [-257, -1.5, -.5, 1.5, 4294967297] : [0, 255, 256, 4660, 65535, 65536, 2147483648, 4294967295, 1.75];
  for (const input of values) {
    const bits = Math.trunc(input) >>> 0;
    const expected = Array.from({length: width}, (_, i) => (bits >>> (8 * (endian === 'Little' ? i : width - i - 1))) & 255);
    const actual = Array.from(run(input), c => c.charCodeAt(0));
    if (JSON.stringify(actual) !== JSON.stringify(expected)) throw Error(`${width}/${endian}/${input}: ${actual} != ${expected}`);
    checks++;
  }
  for (const input of width === 1 ? [NaN, Infinity, -Infinity] : [NaN, Infinity, -Infinity, -1, 4294967296]) {
    let trapped = false;
    try { run(input); } catch (error) { if (!(error instanceof WebAssembly.RuntimeError)) throw error; trapped = true; }
    if (!trapped) throw Error(`Missing Wasm trap ${width}/${endian}/${input}`);
    checks++;
  }
}
console.log(`${checks} native Wasm packing checks passed`);
