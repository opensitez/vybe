const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('key bursts batch without dropping or reordering press and release', async () => {
  const handlers = new Map(), packets = [], acknowledgements = [];
  const document = {body: {}, addEventListener: (kind, handler) => handlers.set(kind, handler)};
  let initialized = false;
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/bridge.js'), 'utf8'), {
    document, queueMicrotask, setTimeout, clearTimeout, cancelAnimationFrame() {}, console,
    fetch: async (url, options) => {
      if (url === 'event') {
        packets.push(JSON.parse(options.body));
        return new Promise(resolve => acknowledgements.push(resolve));
      }
      if (url === 'reply') return {};
      if (initialized) return new Promise(() => {});
      initialized = true;
      return {ok: true, status: 200, headers: {get: () => 'application/json'},
        json: async () => [{id: 1, document: 1, operation: 'NewDocument', args: {title: 'Test'}}]};
    },
  });
  await new Promise(setImmediate);
  const emit = (kind, key) => handlers.get(kind)({target: document.body, key, code: key,
    bubbles: true, composedPath: () => [document.body, document]});
  emit('keydown', 'ArrowUp');
  emit('keydown', 'ArrowDown');
  await new Promise(setImmediate);
  assert.equal(packets.length, 1);
  assert.deepEqual(packets[0].map(event => [event.kind, event.fields.key]),
    [['keydown', 'ArrowUp'], ['keydown', 'ArrowDown']]);
  emit('keyup', 'ArrowUp');
  emit('keyup', 'ArrowDown');
  await new Promise(setImmediate);
  assert.equal(packets.length, 1, 'only one request may be in flight');
  acknowledgements[0]({});
  await new Promise(setImmediate);
  assert.deepEqual(packets[1].map(event => [event.kind, event.fields.key]),
    [['keyup', 'ArrowUp'], ['keyup', 'ArrowDown']]);
  acknowledgements[1]({});
});
