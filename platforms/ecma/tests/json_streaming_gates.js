function check(condition, message) {
    if (!condition) throw new Error(message);
}
function primitives() {
    check(JSON.stringify([true, false, null, undefined, 7, 'line\n"end']) ===
        '[true,false,null,null,7,"line\\n\\"end"]', 'primitive array output');
    check(JSON.stringify({first: true, omitted: undefined, last: 'text'}) ===
        '{"first":true,"last":"text"}', 'omitted member separators');
}
function callbackOrder() {
    const calls = [];
    const custom = {result: 7, toJSON: function(key) { calls.push('toJSON:' + key); return this.result; }};
    const replacer = function(key, value) { calls.push('replace:' + key); return value; };
    check(JSON.stringify({custom: custom, flag: true}, replacer) ===
        '{"custom":7,"flag":true}', 'transformed primitive output');
    check(calls.join('|') === 'replace:|toJSON:custom|replace:custom|replace:flag',
        'toJSON and replacer ordering');
}
function prettyOutput() {
    check(JSON.stringify([1, true], null, 2) === '[\n  1,\n  true\n]', 'array indentation');
    check(JSON.stringify({first: true, omitted: undefined, last: {value: 7}}, null, 2) ===
        '{\n  "first": true,\n  "last": {\n    "value": 7\n  }\n}', 'nested object indentation');
}
const gates = [primitives, callbackOrder, prettyOutput];
let failures = 0;
let index = 0;
while (index < gates.length) {
    try { gates[index](); console.log('JSON streaming gate ' + index + ' passed'); }
    catch (error) { failures++; console.log('JSON streaming gate ' + index + ' failed: ' + error); }
    index++;
}
check(failures === 0, 'JSON streaming failures: ' + failures);
console.log('JSON streaming gates passed');
