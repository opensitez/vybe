/// ES2025+ iterator helpers — map, filter, take, drop, toArray (TC39 proposal)

use super::helpers::run_js;

#[test]
fn iterator_map_helper() {
    assert_eq!(run_js(r#"
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const it = range(5);
// Polyfill using generator
function* mapIter(iter, fn) {
    for (const v of iter) yield fn(v);
}
const doubled = [...mapIter(range(5), x => x * 2)];
console.log(doubled.join(","));
"#), vec!["0,2,4,6,8"]);
}

#[test]
fn iterator_filter_helper() {
    assert_eq!(run_js(r#"
function* filterIter(iter, pred) {
    for (const v of iter) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const evens = [...filterIter(range(10), x => x % 2 === 0)];
console.log(evens.join(","));
"#), vec!["0,2,4,6,8"]);
}

#[test]
fn iterator_take_helper() {
    assert_eq!(run_js(r#"
function* take(n, iter) {
    let count = 0;
    for (const v of iter) {
        if (count++ >= n) break;
        yield v;
    }
}
function* naturals() { let n = 1; while (true) yield n++; }
console.log([...take(5, naturals())].join(","));
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn iterator_drop_helper() {
    assert_eq!(run_js(r#"
function* drop(n, iter) {
    let i = 0;
    for (const v of iter) {
        if (i++ >= n) yield v;
    }
}
const arr = [1, 2, 3, 4, 5];
console.log([...drop(2, arr)].join(","));
"#), vec!["3,4,5"]);
}

#[test]
fn iterator_reduce_helper() {
    assert_eq!(run_js(r#"
function reduceIter(iter, fn, init) {
    let acc = init;
    for (const v of iter) acc = fn(acc, v);
    return acc;
}
function* range(start, end) { for (let i = start; i < end; i++) yield i; }
const sum = reduceIter(range(1, 6), (a, b) => a + b, 0);
console.log(sum);
"#), vec!["15"]);
}

#[test]
fn iterator_flatmap_helper() {
    assert_eq!(run_js(r#"
function* flatMapIter(iter, fn) {
    for (const v of iter) yield* fn(v);
}
const result = [...flatMapIter([[1, 2], [3, 4], [5]], x => x)];
console.log(result.join(","));
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn iterator_zip() {
    assert_eq!(run_js(r#"
function* zipIter(a, b) {
    const iterA = a[Symbol.iterator]();
    const iterB = b[Symbol.iterator]();
    while (true) {
        const rA = iterA.next(), rB = iterB.next();
        if (rA.done || rB.done) break;
        yield [rA.value, rB.value];
    }
}
const result = [...zipIter([1, 2, 3], ["a", "b", "c"])];
console.log(result.map(([a, b]) => a + b).join(","));
"#), vec!["1a,2b,3c"]);
}

#[test]
fn iterator_to_array_pattern() {
    assert_eq!(run_js(r#"
// Polyfill for Iterator.prototype.toArray
function toArray(iter) { return [...iter]; }
function* gen() { yield 1; yield 2; yield 3; }
console.log(toArray(gen()).join(","));
"#), vec!["1,2,3"]);
}

#[test]
fn iterator_some_every() {
    assert_eq!(run_js(r#"
function someIter(iter, pred) {
    for (const v of iter) if (pred(v)) return true;
    return false;
}
function everyIter(iter, pred) {
    for (const v of iter) if (!pred(v)) return false;
    return true;
}
const nums = [2, 4, 6, 7, 8];
console.log(someIter(nums, x => x % 2 !== 0));
console.log(everyIter([2, 4, 6], x => x % 2 === 0));
"#), vec!["true", "true"]);
}

#[test]
fn iterator_pipeline() {
    assert_eq!(run_js(r#"
// Chain operations lazily
function* range(n) { for (let i = 1; i <= n; i++) yield i; }
function* map(fn, iter) { for (const v of iter) yield fn(v); }
function* filter(pred, iter) { for (const v of iter) if (pred(v)) yield v; }
function* take(n, iter) {
    let i = 0;
    for (const v of iter) { if (i++ >= n) break; yield v; }
}
const pipeline = take(3, filter(x => x % 2 === 0, map(x => x * x, range(20))));
console.log([...pipeline].join(","));
"#), vec!["4,16,36"]);
}
