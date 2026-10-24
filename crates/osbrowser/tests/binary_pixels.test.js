const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('binary pixels share the fetched buffer and preserve command order and replies', async () => {
  const commands = [
    { id: 1, document: 1, operation: 'create_element', args: ['canvas'], assigned_node: 7 },
    ...[0, 1].map(index => ({
      id: index + 2, document: 1, operation: 'canvas_image',
      args: { node: 7, command: ['putImageData', { width: 1, height: 1 }, 0, 0] },
      binary_range: [index * 4, 4],
    })),
  ];
  const header = new TextEncoder().encode(JSON.stringify(commands));
  const buffer = new ArrayBuffer(4 + header.length + 8);
  new DataView(buffer).setUint32(0, header.length, true);
  new Uint8Array(buffer, 4, header.length).set(header);
  new Uint8Array(buffer, 4 + header.length).set([1, 2, 3, 255, 4, 5, 6, 255]);
  let served = false;
  const seen = [], replies = [];
  const canvas = { getContext: () => ({}) };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/bridge.js'), 'utf8'), {
    document: { createElement: () => canvas }, DataView, TextDecoder, Uint8Array, Uint8ClampedArray, WeakMap, Map, WeakSet,
    window: { osbrowserCanvas2D: { image: (ctx, command, pixels) => seen.push(pixels) } },
    fetch: async (url, options) => {
      if (url === 'reply') { replies.push(JSON.parse(options.body)); return {}; }
      if (served) throw new Error('finished');
      served = true;
      return { ok: true, status: 200, headers: { get: () => 'application/octet-stream' },
        arrayBuffer: async () => buffer };
    },
    setTimeout: callback => { callback(); return 1; }, clearTimeout() {},
    console: { warn() {}, error() {} },
  });
  await new Promise(setImmediate);
  assert.deepEqual(seen.map(pixels => [...pixels]), [[1, 2, 3, 255], [4, 5, 6, 255]]);
  assert.equal(seen[0].buffer, buffer);
  assert.equal(seen[1].buffer, buffer);
  assert.deepEqual(replies[0].map(reply => reply.id), [1, 2, 3]);
  assert.equal(replies[0][0].result, 7, 'the created node answers under the id the process assigned');
});
