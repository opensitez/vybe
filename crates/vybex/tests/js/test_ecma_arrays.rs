use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Arrays — methods, iteration, modern features
// ═══════════════════════════════════════════════════════════

// ── Creation ───────────────────────────────────────────────

#[test]
fn array_literal() {
    let out = run_js(r#"
const arr = [1, 2, 3];
console.log(arr.length);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_from() {
    let out = run_js(r#"
const arr = Array.from([1, 2, 3]);
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn array_isarray() {
    let out = run_js(r#"
console.log(Array.isArray([1, 2]));
console.log(Array.isArray("hello"));
console.log(Array.isArray({ length: 1 }));
"#);
    assert_eq!(out, vec!["true", "false", "false"]);
}

// ── Mutating methods ───────────────────────────────────────

#[test]
fn push_pop() {
    let out = run_js(r#"
const arr = [1, 2];
arr.push(3);
console.log(arr.length);
const last = arr.pop();
console.log(last);
console.log(arr.length);
"#);
    assert_eq!(out, vec!["3", "3", "2"]);
}

#[ignore]
#[test]
fn shift_unshift() {
    let out = run_js(r#"
const arr = [2, 3];
arr.unshift(1);
console.log(arr.join(","));
const first = arr.shift();
console.log(first);
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,2,3", "1", "2,3"]);
}

#[test]
fn splice_remove() {
    let out = run_js(r#"
const arr = [1, 2, 3, 4, 5];
arr.splice(1, 2);
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,4,5"]);
}

#[test]
fn splice_insert() {
    let out = run_js(r#"
const arr = [1, 4, 5];
arr.splice(1, 0, 2, 3);
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,2,3,4,5"]);
}

// ── Non-mutating methods ───────────────────────────────────

#[test]
fn concat() {
    let out = run_js(r#"
const a = [1, 2];
const b = [3, 4];
console.log(a.concat(b).join(","));
"#);
    assert_eq!(out, vec!["1,2,3,4"]);
}

#[test]
fn slice() {
    let out = run_js(r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.slice(1, 3).join(","));
console.log(arr.slice(3).join(","));
"#);
    assert_eq!(out, vec!["2,3", "4,5"]);
}

#[test]
fn join() {
    let out = run_js(r#"
console.log([1, 2, 3].join("-"));
console.log(["a", "b", "c"].join(""));
"#);
    assert_eq!(out, vec!["1-2-3", "abc"]);
}

#[test]
fn reverse() {
    let out = run_js(r#"
const arr = [1, 2, 3];
arr.reverse();
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["3,2,1"]);
}

#[test]
fn sort_default() {
    let out = run_js(r#"
const arr = [3, 1, 4, 1, 5];
arr.sort();
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,1,3,4,5"]);
}

#[test]
fn sort_with_comparator() {
    let out = run_js(r#"
const arr = [3, 1, 4, 1, 5];
arr.sort((a, b) => b - a);
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["5,4,3,1,1"]);
}

#[test]
fn flat() {
    let out = run_js(r#"
const arr = [[1, 2], [3, 4], [5]];
console.log(arr.flat().join(","));
"#);
    assert_eq!(out, vec!["1,2,3,4,5"]);
}

#[test]
fn fill() {
    let out = run_js(r#"
const arr = [1, 2, 3, 4, 5];
arr.fill(0, 1, 4);
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,0,0,0,5"]);
}

// ── Search methods ─────────────────────────────────────────

#[ignore]
#[test]
fn indexof() {
    let out = run_js(r#"
const arr = [10, 20, 30, 20];
console.log(arr.indexOf(20));
console.log(arr.indexOf(40));
"#);
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn includes() {
    let out = run_js(r#"
console.log([1, 2, 3].includes(2));
console.log([1, 2, 3].includes(4));
"#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn find() {
    let out = run_js(r#"
const arr = [1, 2, 3, 4, 5];
const found = arr.find(x => x > 3);
console.log(found);
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn findindex() {
    let out = run_js(r#"
const arr = [1, 2, 3, 4, 5];
const idx = arr.findIndex(x => x > 3);
console.log(idx);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn some() {
    let out = run_js(r#"
console.log([1, 2, 3].some(x => x > 2));
console.log([1, 2, 3].some(x => x > 5));
"#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn every() {
    let out = run_js(r#"
console.log([2, 4, 6].every(x => x % 2 === 0));
console.log([2, 3, 6].every(x => x % 2 === 0));
"#);
    assert_eq!(out, vec!["true", "false"]);
}

// ── Iteration methods ──────────────────────────────────────

#[test]
fn map() {
    let out = run_js(r#"
const arr = [1, 2, 3];
const doubled = arr.map(x => x * 2);
console.log(doubled.join(","));
"#);
    assert_eq!(out, vec!["2,4,6"]);
}

#[test]
fn filter() {
    let out = run_js(r#"
const arr = [1, 2, 3, 4, 5, 6];
const evens = arr.filter(x => x % 2 === 0);
console.log(evens.join(","));
"#);
    assert_eq!(out, vec!["2,4,6"]);
}

#[test]
fn reduce() {
    let out = run_js(r#"
const sum = [1, 2, 3, 4].reduce((acc, x) => acc + x, 0);
console.log(sum);
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn foreach() {
    let out = run_js(r#"
let sum = 0;
[1, 2, 3].forEach(x => { sum += x; });
console.log(sum);
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn map_filter_chain() {
    let out = run_js(r#"
const result = [1, 2, 3, 4, 5]
    .filter(x => x % 2 !== 0)
    .map(x => x * x);
console.log(result.join(","));
"#);
    assert_eq!(out, vec!["1,9,25"]);
}

#[ignore]
#[test]
fn reduce_to_object() {
    let out = run_js(r#"
const pairs = [["a", 1], ["b", 2], ["c", 3]];
const obj = pairs.reduce((acc, [k, v]) => {
    acc[k] = v;
    return acc;
}, {});
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

// ── Modern array methods ───────────────────────────────────

#[ignore]
#[test]
fn array_at() {
    let out = run_js(r#"
const arr = [10, 20, 30, 40, 50];
console.log(arr.at(0));
console.log(arr.at(-1));
console.log(arr.at(-2));
"#);
    assert_eq!(out, vec!["10", "50", "40"]);
}

#[test]
fn array_flat_nested() {
    let out = run_js(r#"
const arr = [1, [2, [3, [4]]]];
console.log(arr.flat().join(","));
"#);
    assert_eq!(out, vec!["1,2,3,4"]);
}

#[test]
fn array_entries_pattern() {
    let out = run_js(r#"
const arr = ["a", "b", "c"];
let result = [];
for (let i = 0; i < arr.length; i++) {
    result.push(i + ":" + arr[i]);
}
console.log(result.join(","));
"#);
    assert_eq!(out, vec!["0:a,1:b,2:c"]);
}
