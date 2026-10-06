const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('bridge stops polling after its local server disappears', async () => {
  const source = fs.readFileSync(path.join(__dirname, '../src/bridge.js'), 'utf8');
  let requests = 0;
  const warnings = [];
  vm.runInNewContext(source, {
    document: {},
    fetch: async () => {
      requests++;
      throw new Error('connection refused');
    },
    setTimeout: callback => { callback(); return 1; },
    clearTimeout: () => {},
    cancelAnimationFrame: () => {},
    console: { warn: message => warnings.push(message) },
  });
  await new Promise(setImmediate);
  assert.equal(requests, 3);
  assert.equal(warnings.length, 1);
});
