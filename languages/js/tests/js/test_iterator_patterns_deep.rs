/// Iterator protocol and iteration patterns
use super::helpers::run_js;

#[test]
fn manual_iterator_usage() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, 20, 30];
const iter = arr[Symbol.iterator]();
console.log(iter.next().value);
console.log(iter.next().value);
console.log(iter.next().done);
console.log(iter.next().done);
"#
        ),
        vec!["10", "20", "false", "true"]
    );
}

#[test]
fn string_iteration_unicode() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
const chars = [...s];
console.log(chars.length);
console.log(chars[0]);
// for-of iterates code points
let count = 0;
for (const c of "abc") count++;
console.log(count);
"#
        ),
        vec!["5", "h", "3"]
    );
}

#[test]
fn map_set_iteration() {
    assert_eq!(
        run_js(
            r#"
const map = new Map([["a", 1], ["b", 2]]);
const keyIter = map.keys();
console.log(keyIter.next().value);
const set = new Set([10, 20, 30]);
const setIter = set[Symbol.iterator]();
console.log(setIter.next().value);
console.log([...set.values()].join(","));
"#
        ),
        vec!["a", "10", "10,20,30"]
    );
}

#[test]
fn entries_iterator_pattern() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const entries = Object.entries(obj);
for (const [key, val] of entries) {
    if (key === "b") { console.log(val); break; }
}
// Array entries
const arr = ["x", "y", "z"];
for (const [i, v] of arr.entries()) {
    if (i === 1) { console.log(v); break; }
}
"#
        ),
        vec!["2", "y"]
    );
}

#[test]
fn infinite_iterator_take() {
    assert_eq!(
        run_js(
            r#"
function* naturals() { let n = 1; while (true) yield n++; }
function take(n, iter) {
    const result = [];
    for (const v of iter) { result.push(v); if (result.length >= n) break; }
    return result;
}
console.log(take(5, naturals()).join(","));
console.log(take(3, naturals()).join(","));
"#
        ),
        vec!["1,2,3,4,5", "1,2,3"]
    );
}

#[test]
fn generator_return_and_throw() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        yield "cleanup";
    }
}
const g1 = gen();
console.log(g1.next().value);
const ret = g1.return("early");
console.log(ret.value);
console.log(g1.next().done);
"#
        ),
        vec!["1", "cleanup", "true"]
    );
}

#[test]
fn iterator_to_array_methods() {
    assert_eq!(
        run_js(
            r#"
function* range(start, end) {
    for (let i = start; i <= end; i++) yield i;
}
// Array.from accepts iterables
const arr = Array.from(range(1, 5));
console.log(arr.join(","));
// Spread also works
const doubled = [...range(1, 3)].map(x => x * 2);
console.log(doubled.join(","));
"#
        ),
        vec!["1,2,3,4,5", "2,4,6"]
    );
}

#[test]
fn iterable_destructuring_partial() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 10; yield 20; yield 30; yield 40; }
const [a, b, ...rest] = gen();
console.log(a);
console.log(b);
console.log(rest.join(","));
"#
        ),
        vec!["10", "20", "30,40"]
    );
}

#[test]
fn for_of_with_index_via_entries() {
    assert_eq!(
        run_js(
            r#"
const items = ["a", "b", "c", "d"];
const indexed = [];
for (const [i, v] of items.entries()) indexed.push(`${i}:${v}`);
console.log(indexed.join(","));
"#
        ),
        vec!["0:a,1:b,2:c,3:d"]
    );
}

#[test]
fn generator_composition_pipeline() {
    assert_eq!(
        run_js(
            r#"
function* map(iter, fn) { for (const v of iter) yield fn(v); }
function* filter(iter, pred) { for (const v of iter) if (pred(v)) yield v; }
function* take(iter, n) { let i=0; for (const v of iter) { if (i++>=n) break; yield v; } }
function* counter(start=0) { while(true) yield start++; }

const result = [...take(filter(map(counter(), x=>x*x), x=>x%2===0), 4)];
console.log(result.join(","));
"#
        ),
        vec!["0,4,16,36"]
    );
}

#[test]
fn custom_iterable_class() {
    assert_eq!(
        run_js(
            r#"
class Matrix {
    constructor(rows) { this.rows = rows; }
    [Symbol.iterator]() {
        let r = 0, c = 0;
        const rows = this.rows;
        return {
            next() {
                if (r >= rows.length) return { done: true };
                const value = rows[r][c++];
                if (c >= rows[r].length) { c = 0; r++; }
                return { value, done: false };
            }
        };
    }
}
const m = new Matrix([[1,2],[3,4],[5,6]]);
console.log([...m].join(","));
"#
        ),
        vec!["1,2,3,4,5,6"]
    );
}

#[test]
fn lazy_sequence_operations() {
    assert_eq!(
        run_js(
            r#"
function* seq(arr) { yield* arr; }
function* mapLazy(gen, fn) { for (const v of gen) yield fn(v); }
function reduce(gen, fn, init) {
    let acc = init;
    for (const v of gen) acc = fn(acc, v);
    return acc;
}
const sum = reduce(mapLazy(seq([1,2,3,4,5]), x => x*x), (a,b) => a+b, 0);
console.log(sum);
"#
        ),
        vec!["55"]
    );
}
