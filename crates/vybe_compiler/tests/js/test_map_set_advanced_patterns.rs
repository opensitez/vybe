/// Map and Set advanced patterns — fromEntries, groupBy, set algebra

use super::helpers::run_js;

#[test]
fn map_from_entries() {
    assert_eq!(run_js(r#"
const m = new Map(Object.entries({ a: 1, b: 2, c: 3 }));
console.log(m.get("a"));
console.log(m.get("c"));
console.log(m.size);
"#), vec!["1", "3", "3"]);
}

#[test]
fn map_transform_then_back_to_object() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2, c: 3 };
const doubled = Object.fromEntries(
    Object.entries(obj).map(([k, v]) => [k, v * 2])
);
console.log(doubled.a);
console.log(doubled.b);
console.log(doubled.c);
"#), vec!["2", "4", "6"]);
}

#[test]
fn set_union() {
    assert_eq!(run_js(r#"
const a = new Set([1, 2, 3]);
const b = new Set([2, 3, 4, 5]);
const union = new Set([...a, ...b]);
console.log([...union].sort((a,b)=>a-b).join(","));
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn set_intersection() {
    assert_eq!(run_js(r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([2, 4, 6]);
const intersection = new Set([...a].filter(x => b.has(x)));
console.log([...intersection].sort((a,b)=>a-b).join(","));
"#), vec!["2,4"]);
}

#[test]
fn set_difference() {
    assert_eq!(run_js(r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([2, 4]);
const diff = new Set([...a].filter(x => !b.has(x)));
console.log([...diff].sort((a,b)=>a-b).join(","));
"#), vec!["1,3"]);
}

#[test]
fn map_chaining_operations() {
    assert_eq!(run_js(r#"
const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
const result = [...m.entries()]
    .filter(([, v]) => v > 1)
    .map(([k, v]) => k + "=" + v);
console.log(result.join(","));
"#), vec!["b=2,c=3"]);
}

#[test]
fn set_nan_deduplication() {
    assert_eq!(run_js(r#"
const s = new Set([NaN, NaN, NaN]);
console.log(s.size);
console.log(s.has(NaN));
"#), vec!["1", "true"]);
}

#[test]
fn map_with_complex_keys() {
    assert_eq!(run_js(r#"
const keyA = { id: 1 };
const keyB = { id: 2 };
const m = new Map();
m.set(keyA, "first");
m.set(keyB, "second");
m.set({id:1}, "third"); // different object reference
console.log(m.size);
console.log(m.get(keyA));
"#), vec!["3", "first"]);
}

#[test]
fn object_groupby_pattern() {
    assert_eq!(run_js(r#"
const people = [
    { name: "Alice", dept: "eng" },
    { name: "Bob", dept: "hr" },
    { name: "Charlie", dept: "eng" },
];
const grouped = Object.groupBy(people, p => p.dept);
console.log(grouped.eng.length);
console.log(grouped.hr.length);
console.log(grouped.eng[0].name);
"#), vec!["2", "1", "Alice"]);
}

#[test]
fn map_groupby_pattern() {
    assert_eq!(run_js(r#"
const items = [1, 2, 3, 4, 5, 6];
const grouped = Map.groupBy(items, x => x % 2 === 0 ? "even" : "odd");
console.log(grouped.get("even").join(","));
console.log(grouped.get("odd").join(","));
"#), vec!["2,4,6", "1,3,5"]);
}

#[test]
fn set_has_and_iteration_order() {
    assert_eq!(run_js(r#"
const s = new Set();
s.add("c").add("a").add("b");
console.log([...s].join(","));
console.log(s.has("a"));
s.delete("a");
console.log([...s].join(","));
"#), vec!["c,a,b", "true", "c,b"]);
}
