const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('canvas source drawing stays native, including self-copy and crop arguments', () => {
  const calls = [];
  class Canvas {
    getContext() {
      return {
        drawImage: (...args) => calls.push(args),
        getImageData: () => { throw new Error('unexpected pixel readback'); },
        putImageData: () => { throw new Error('unexpected pixel upload'); },
      };
    }
  }
  const destination = new Canvas();
  const source = new Canvas();
  const nodes = new Map([[1, destination], [2, source]]);
  const sandbox = { window: {}, HTMLCanvasElement: Canvas };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/canvas.js'), 'utf8'), sandbox);
  const draw = sandbox.window.vybeCanvas;
  const rect = [2, 3, 10, 11, 12, 13, 20, 22];
  assert.equal(draw('CanvasDrawImage', {
    target: 'n1', payload: { source: 'n2', rect },
  }, {}, nodes), 'None');
  assert.deepEqual(calls[0], [source, ...rect]);
  assert.equal(draw('CanvasDrawImage', {
    target: 'n1', payload: { source: 'n1', rect },
  }, {}, nodes), 'None');
  assert.deepEqual(calls[1], [destination, ...rect]);
  assert.equal(draw('CanvasDrawImage', {
    target: 'n1', payload: { source: 'n3', rect },
  }, {}, nodes), 'Absent');
  assert.equal(calls.length, 2);
});

test('native drawing leaves browser errors intact rather than falling back to pixel readback', () => {
  class Canvas {
    getContext() {
      return { drawImage: () => { throw new Error('InvalidStateError'); } };
    }
  }
  const sandbox = { window: {}, HTMLCanvasElement: Canvas };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/canvas.js'), 'utf8'), sandbox);
  assert.throws(() => sandbox.window.vybeCanvas('CanvasDrawImage', {
    target: 'n1', payload: { source: 'n2', rect: [0, 0, 1, 1, 0, 0, 1, 1] },
  }, {}, new Map([[1, new Canvas()], [2, new Canvas()]])), /InvalidStateError/);
});
