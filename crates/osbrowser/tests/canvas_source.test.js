const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// A 2D context that records calls and serialises `fillStyle`/`font` the way a
// browser does: an invalid value leaves the previous one in place.
function context(calls = []) {
  let fill = '#000000', font = '10px sans-serif';
  return new Proxy({}, {
    get: (_, name) => {
      if (name === 'fillStyle') return fill;
      if (name === 'font') return font;
      return (...args) => calls.push([name, ...args]);
    },
    set: (_, name, value) => {
      if (name === 'fillStyle') { if (/^(#|rgb)/.test(value)) fill = value === 'rgb(1, 2, 3)' ? '#010203' : value; }
      else if (name === 'font') { if (/\d+px/.test(value)) font = value; }
      else calls.push(['=' + name, value]);
      return true;
    },
  });
}

function load() {
  const sandbox = { window: {}, document: { createElement: () => ({ getContext: () => context() }) } };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/canvas.js'), 'utf8'), sandbox);
  return sandbox.window.osbrowserCanvas2D;
}

test('a recorded transcript reaches the real context in order', () => {
  const calls = [];
  const canvas2d = load();
  canvas2d.draw(context(calls), {}, [
    ['save'], ['set', 'lineWidth', 3], ['beginPath'], ['moveTo', 1, 2], ['lineTo', 3, 4],
    ['stroke'], ['fillRect', 0, 0, 5, 6], ['noop'], ['restore'],
  ]);
  assert.deepEqual(calls, [
    ['save'], ['=lineWidth', 3], ['beginPath'], ['moveTo', 1, 2], ['lineTo', 3, 4],
    ['stroke'], ['fillRect', 0, 0, 5, 6], ['restore'],
  ]);
});

test('css parsing tells an invalid value from one equal to the sentinel', () => {
  const canvas2d = load();
  assert.deepEqual([...canvas2d.parseColor('rgb(1, 2, 3)')], [1, 2, 3, 255]);
  assert.equal(canvas2d.parseColor('not a colour'), null);
  assert.equal(canvas2d.parseFont('not a font'), null);
  assert.deepEqual({ ...canvas2d.parseFont('italic bold 14px serif') },
    { family: 'serif', size: 14, bold: true, italic: true });
});
