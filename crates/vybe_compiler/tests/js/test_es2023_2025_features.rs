/// ES2023-2025 new array/object/string methods

use super::helpers::run_js;

#[test]
fn array_to_sorted_non_mutating() {
    assert_eq!(run_js(r#"
const orig = [3, 1, 2];
const sorted = orig.toSorted((a, b) => a - b);
console.log(sorted.join(","));
console.log(orig.join(","));
"#), vec!["1,2,3", "3,1,2"]);
}

#[test]
fn array_to_reversed_non_mutating() {
    assert_eq!(run_js(r#"
const orig = [1, 2, 3];
const rev = orig.toReversed();
console.log(rev.join(","));
console.log(orig.join(","));
"#), vec!["3,2,1", "1,2,3"]);
}

#[test]
fn array_to_spliced_non_mutating() {
    assert_eq!(run_js(r#"
const orig = [1, 2, 3, 4, 5];
const spliced = orig.toSpliced(2, 1, 99);
console.log(spliced.join(","));
console.log(orig.join(","));
"#), vec!["1,2,99,4,5", "1,2,3,4,5"]);
}

#[test]
fn array_with_non_mutating() {
    assert_eq!(run_js(r#"
const orig = [1, 2, 3];
const updated = orig.with(1, 99);
console.log(updated.join(","));
console.log(orig.join(","));
"#), vec!["1,99,3", "1,2,3"]);
}

#[test]
fn array_find_last() {
    assert_eq!(run_js(r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.findLast(x => x % 2 === 0));
console.log(arr.findLastIndex(x => x % 2 === 0));
"#), vec!["4", "3"]);
}

#[test]
fn object_has_own() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
console.log(Object.hasOwn(obj, "a"));
console.log(Object.hasOwn(obj, "toString"));
const n = Object.create(null);
n.x = 5;
console.log(Object.hasOwn(n, "x"));
"#), vec!["true", "false", "true"]);
}

#[test]
fn error_cause_property() {
    assert_eq!(run_js(r#"
function fetchData() {
    try { throw new Error("network failure"); }
    catch (e) { throw new Error("Failed to fetch", { cause: e }); }
}
try {
    fetchData();
} catch (e) {
    console.log(e.message);
    console.log(e.cause.message);
}
"#), vec!["Failed to fetch", "network failure"]);
}

#[test]
fn at_method_strings() {
    assert_eq!(run_js(r#"
const s = "hello";
console.log(s.at(0));
console.log(s.at(-1));
console.log(s.at(-2));
console.log(s.at(10));
"#), vec!["h", "o", "l", "undefined"]);
}

#[test]
fn string_replace_all() {
    assert_eq!(run_js(r#"
const s = "foo bar foo baz foo";
console.log(s.replaceAll("foo", "qux"));
console.log("a.b.c".replaceAll(".", "-"));
"#), vec!["qux bar qux baz qux", "a-b-c"]);
}

#[test]
fn promise_any_first_resolve() {
    assert_eq!(run_js(r#"
async function test() {
    const p = await Promise.any([
        Promise.reject(1),
        Promise.resolve(2),
        Promise.resolve(3)
    ]);
    console.log(p);
}
test();
"#), vec!["2"]);
}

#[test]
fn structuredclone_deep_copy() {
    assert_eq!(run_js(r#"
const orig = { a: 1, b: { c: [1, 2, 3] } };
const copy = structuredClone(orig);
copy.b.c.push(4);
console.log(orig.b.c.length);
console.log(copy.b.c.length);
"#), vec!["3", "4"]);
}

#[test]
fn array_group_by_object() {
    assert_eq!(run_js(r#"
const items = [
    { type: "A", val: 1 }, { type: "B", val: 2 }, { type: "A", val: 3 }
];
const grouped = Object.groupBy(items, item => item.type);
console.log(grouped.A.length);
console.log(grouped.B.length);
console.log(grouped.A[0].val);
"#), vec!["2", "1", "1"]);
}

#[test]
fn map_group_by() {
    assert_eq!(run_js(r#"
const words = ["apple", "banana", "avocado", "blueberry"];
const grouped = Map.groupBy(words, w => w[0]);
console.log(grouped.get("a").join(","));
console.log(grouped.get("b").join(","));
"#), vec!["apple,avocado", "banana,blueberry"]);
}
