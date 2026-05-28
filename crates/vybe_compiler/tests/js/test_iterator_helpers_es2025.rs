/// Iterator helper methods (ES2025 Stage 3) — Iterator.prototype.map, filter,
/// take, drop, flatMap, reduce, toArray, forEach, some, every, find, from.

use super::helpers::run_js;

// ── Iterator.from ─────────────────────────────────────────────────────────────

#[test]
fn iterator_from_array() {
    assert_eq!(run_js(r#"
const iter = Iterator.from([1, 2, 3]);
const result = iter.toArray();
console.log(result.join(","));
"#), vec!["1,2,3"]);
}

#[test]
fn iterator_from_generator() {
    assert_eq!(run_js(r#"
function* gen() { yield 1; yield 2; yield 3; }
const result = Iterator.from(gen()).toArray();
console.log(result.join(","));
"#), vec!["1,2,3"]);
}

#[test]
fn iterator_from_string() {
    assert_eq!(run_js(r#"
const result = Iterator.from("abc").toArray();
console.log(result.join(","));
"#), vec!["a,b,c"]);
}

// ── Iterator.prototype.map ────────────────────────────────────────────────────

#[test]
fn iterator_map_transforms_values() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3]).map(x => x * 2).toArray();
console.log(result.join(","));
"#), vec!["2,4,6"]);
}

#[test]
fn iterator_map_is_lazy() {
    assert_eq!(run_js(r#"
const called = [];
const iter = Iterator.from([1, 2, 3]).map(x => { called.push(x); return x * 2; });
console.log(called.length);
iter.next();
console.log(called.length);
"#), vec!["0", "1"]);
}

#[test]
fn iterator_map_chained() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3])
    .map(x => x + 1)
    .map(x => x * 2)
    .toArray();
console.log(result.join(","));
"#), vec!["4,6,8"]);
}

// ── Iterator.prototype.filter ─────────────────────────────────────────────────

#[test]
fn iterator_filter_keeps_matching() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4, 5]).filter(x => x % 2 === 0).toArray();
console.log(result.join(","));
"#), vec!["2,4"]);
}

#[test]
fn iterator_filter_and_map_chain() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4, 5])
    .filter(x => x % 2 !== 0)
    .map(x => x ** 2)
    .toArray();
console.log(result.join(","));
"#), vec!["1,9,25"]);
}

// ── Iterator.prototype.take ───────────────────────────────────────────────────

#[test]
fn iterator_take_limits_count() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4, 5]).take(3).toArray();
console.log(result.join(","));
"#), vec!["1,2,3"]);
}

#[test]
fn iterator_take_from_infinite() {
    assert_eq!(run_js(r#"
function* naturals() { let n = 1; while (true) yield n++; }
const result = Iterator.from(naturals()).take(5).toArray();
console.log(result.join(","));
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn iterator_take_zero_returns_empty() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3]).take(0).toArray();
console.log(result.length);
"#), vec!["0"]);
}

// ── Iterator.prototype.drop ───────────────────────────────────────────────────

#[test]
fn iterator_drop_skips_first_n() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4, 5]).drop(2).toArray();
console.log(result.join(","));
"#), vec!["3,4,5"]);
}

#[test]
fn iterator_drop_and_take_combo() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4, 5, 6]).drop(2).take(3).toArray();
console.log(result.join(","));
"#), vec!["3,4,5"]);
}

// ── Iterator.prototype.flatMap ────────────────────────────────────────────────

#[test]
fn iterator_flatmap_flattens_one_level() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3]).flatMap(x => [x, x * 10]).toArray();
console.log(result.join(","));
"#), vec!["1,10,2,20,3,30"]);
}

#[test]
fn iterator_flatmap_with_generator() {
    assert_eq!(run_js(r#"
const result = Iterator.from([2, 3]).flatMap(function*(x) {
    for (let i = 1; i <= x; i++) yield i;
}).toArray();
console.log(result.join(","));
"#), vec!["1,2,1,2,3"]);
}

// ── Iterator.prototype.reduce ─────────────────────────────────────────────────

#[test]
fn iterator_reduce_sum() {
    assert_eq!(run_js(r#"
const sum = Iterator.from([1, 2, 3, 4, 5]).reduce((acc, x) => acc + x, 0);
console.log(sum);
"#), vec!["15"]);
}

#[test]
fn iterator_reduce_without_initial() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4]).reduce((acc, x) => acc + x);
console.log(result);
"#), vec!["10"]);
}

// ── Iterator.prototype.toArray ────────────────────────────────────────────────

#[test]
fn iterator_toarray_collects_all() {
    assert_eq!(run_js(r#"
const arr = Iterator.from([1, 2, 3]).toArray();
console.log(Array.isArray(arr));
console.log(arr.length);
"#), vec!["true", "3"]);
}

// ── Iterator.prototype.forEach ────────────────────────────────────────────────

#[test]
fn iterator_foreach_visits_all() {
    assert_eq!(run_js(r#"
const seen = [];
Iterator.from([10, 20, 30]).forEach(x => seen.push(x));
console.log(seen.join(","));
"#), vec!["10,20,30"]);
}

// ── Iterator.prototype.some / every ──────────────────────────────────────────

#[test]
fn iterator_some_returns_true_on_match() {
    assert_eq!(run_js(r#"
console.log(Iterator.from([1, 2, 3]).some(x => x > 2));
console.log(Iterator.from([1, 2, 3]).some(x => x > 10));
"#), vec!["true", "false"]);
}

#[test]
fn iterator_every_returns_true_when_all_match() {
    assert_eq!(run_js(r#"
console.log(Iterator.from([2, 4, 6]).every(x => x % 2 === 0));
console.log(Iterator.from([2, 3, 6]).every(x => x % 2 === 0));
"#), vec!["true", "false"]);
}

// ── Iterator.prototype.find ───────────────────────────────────────────────────

#[test]
fn iterator_find_first_match() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3, 4]).find(x => x > 2);
console.log(result);
"#), vec!["3"]);
}

#[test]
fn iterator_find_no_match_returns_undefined() {
    assert_eq!(run_js(r#"
const result = Iterator.from([1, 2, 3]).find(x => x > 10);
console.log(result);
"#), vec!["undefined"]);
}
