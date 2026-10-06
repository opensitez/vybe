function check(condition, message) {
    if (!condition) throw new Error(message);
}
const receiver = {base: 10};
function fixed(a, b, c, d) { return this.base + a + b + c + d; }
function rest(first, ...remaining) { return this.base + first + remaining.length; }
function wideRest(a, b, c, d, e, f, g, h, i, ...remaining) {
    return this.base + a + i + remaining.length;
}
const four = [1, 2, 3, 4];
const many = [];
let index = 0;
while (index < 40) { many.push(index + 1); index++; }
check(fixed.call(receiver, 1, 2, 3, 4) === 20, 'fixed call receiver and arguments');
check(fixed.apply(receiver, four) === 20, 'fixed apply receiver and arguments');
check(Reflect.apply(fixed, receiver, four) === 20, 'fixed Reflect.apply');
check(rest.call(receiver, 1) === 11, 'empty rest');
check(rest.apply(receiver, four) === 14, 'short rest');
check(Reflect.apply(rest, receiver, many) === 50, 'long rest');
check(Reflect.apply(wideRest, receiver, many) === 51, 'heap-packed fixed and rest arguments');
check(rest.bind(receiver, 1).apply(null, [2, 3]) === 13, 'bound rest');
const sentinel = {};
function throwing(...values) { if (values.length === 4) throw sentinel; }
let caught = false;
try { Reflect.apply(throwing, receiver, four); } catch (error) { caught = error === sentinel; }
check(caught, 'rest call preserves thrown identity');
console.log('call metadata gates passed');
function benchmark(fn, label, expected) {
    let warm = 0;
    while (warm < 100) { fn.apply(receiver, four); warm++; }
    let iteration = 0;
    let result = 0;
    const start = Date.now();
    while (iteration < 10000) { result = fn.apply(receiver, four); iteration++; }
    const elapsed = Date.now() - start;
    check(result === expected, label + ' benchmark result');
    console.log(label + ': ' + elapsed + ' ms / 10000 calls');
}
benchmark(fixed, 'fixed apply', 20);
benchmark(rest, 'rest apply', 14);
