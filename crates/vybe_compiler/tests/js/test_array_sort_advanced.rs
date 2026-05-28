/// Array sorting — stability, custom comparators, edge cases with NaN/undefined,
/// sort with types, sort mutation, Schwartzian transform.

use super::helpers::run_js;

// ── basic sort stability ──────────────────────────────────────────────────────

#[test]
fn sort_is_stable() {
    assert_eq!(run_js(r#"
// Stable sort preserves relative order of equal elements
const items = [
    { priority: 1, name: "c" },
    { priority: 2, name: "a" },
    { priority: 1, name: "b" },
];
items.sort((a, b) => a.priority - b.priority);
// Equal priority items (c and b) must retain c before b
console.log(items[0].name);
console.log(items[1].name);
"#), vec!["c", "b"]);
}

#[test]
fn sort_default_converts_to_string() {
    assert_eq!(run_js(r#"
const arr = [10, 9, 2, 21, 3];
arr.sort();
// Default sort is lexicographic
console.log(arr.join(","));
"#), vec!["10,2,21,3,9"]);
}

#[test]
fn sort_numeric_comparator() {
    assert_eq!(run_js(r#"
const arr = [10, 9, 2, 21, 3];
arr.sort((a, b) => a - b);
console.log(arr.join(","));
"#), vec!["2,3,9,10,21"]);
}

#[test]
fn sort_descending() {
    assert_eq!(run_js(r#"
const arr = [5, 2, 8, 1, 9];
arr.sort((a, b) => b - a);
console.log(arr.join(","));
"#), vec!["9,8,5,2,1"]);
}

// ── sort with undefined ───────────────────────────────────────────────────────

#[test]
fn sort_undefined_moves_to_end() {
    assert_eq!(run_js(r#"
const arr = [3, undefined, 1, undefined, 2];
arr.sort();
// undefined values go to the end
console.log(arr[arr.length - 1] === undefined);
console.log(arr[arr.length - 2] === undefined);
"#), vec!["true", "true"]);
}

// ── sort mutates original ─────────────────────────────────────────────────────

#[test]
fn sort_mutates_in_place() {
    assert_eq!(run_js(r#"
const arr = [3, 1, 2];
const ret = arr.sort();
console.log(ret === arr); // same array
console.log(arr.join(","));
"#), vec!["true", "1,2,3"]);
}

// ── sort strings ──────────────────────────────────────────────────────────────

#[test]
fn sort_strings_alphabetically() {
    assert_eq!(run_js(r#"
const words = ["banana", "apple", "cherry", "date"];
words.sort();
console.log(words.join(","));
"#), vec!["apple,banana,cherry,date"]);
}

#[test]
fn sort_strings_case_sensitive() {
    assert_eq!(run_js(r#"
const words = ["banana", "Apple", "cherry"];
words.sort();
// uppercase comes before lowercase in Unicode
console.log(words[0]);
"#), vec!["Apple"]);
}

// ── sort objects by property ──────────────────────────────────────────────────

#[test]
fn sort_objects_by_date() {
    assert_eq!(run_js(r#"
const events = [
    { name: "c", date: new Date(2024, 2, 1) },
    { name: "a", date: new Date(2024, 0, 1) },
    { name: "b", date: new Date(2024, 1, 1) },
];
events.sort((a, b) => a.date - b.date);
console.log(events.map(e => e.name).join(","));
"#), vec!["a,b,c"]);
}

// ── Schwartzian transform (sort with computed key) ────────────────────────────

#[test]
fn schwartzian_transform_sort_by_computed_key() {
    assert_eq!(run_js(r#"
const words = ["banana", "fig", "cherry", "apple"];
const sorted = words
    .map(w => [w, w.length])
    .sort((a, b) => a[1] - b[1])
    .map(([w]) => w);
console.log(sorted.join(","));
"#), vec!["fig,apple,banana,cherry"]);
}

// ── sort empty and single element ────────────────────────────────────────────

#[test]
fn sort_empty_array() {
    assert_eq!(run_js(r#"
const arr = [];
arr.sort();
console.log(arr.length);
"#), vec!["0"]);
}

#[test]
fn sort_single_element_unchanged() {
    assert_eq!(run_js(r#"
const arr = [42];
arr.sort();
console.log(arr[0]);
"#), vec!["42"]);
}

// ── sort with localeCompare ────────────────────────────────────────────────────

#[test]
fn sort_locale_sensitive() {
    assert_eq!(run_js(r#"
const words = ["résumé", "apple", "éclair"];
words.sort((a, b) => a.localeCompare(b));
// All should be sorted — exact order depends on locale but shouldn't throw
console.log(words.length);
"#), vec!["3"]);
}

// ── sort with mixed types via toString ────────────────────────────────────────

#[test]
fn sort_mixed_types_coerced_to_string() {
    assert_eq!(run_js(r#"
const arr = [null, 1, "a"];
arr.sort();
// null → "null", 1 → "1", "a" → "a"
// lexicographic: "1" < "a" < "null"
console.log(arr[0]);
"#), vec!["1"]);
}

// ── sort stability with complex key ──────────────────────────────────────────

#[test]
fn sort_stable_complex_key() {
    assert_eq!(run_js(r#"
const data = [
    { key: "b", order: 0 },
    { key: "a", order: 1 },
    { key: "b", order: 2 },
    { key: "a", order: 3 },
];
data.sort((a, b) => a.key.localeCompare(b.key));
// Stable: a's keep order 1,3; b's keep order 0,2
console.log(data[0].key + data[0].order);
console.log(data[1].key + data[1].order);
console.log(data[2].key + data[2].order);
console.log(data[3].key + data[3].order);
"#), vec!["a1", "a3", "b0", "b2"]);
}
