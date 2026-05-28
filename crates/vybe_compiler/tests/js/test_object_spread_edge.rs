/// Object spread — own enumerable only, order, nested, merging, getters

use super::helpers::run_js;

#[test]
fn spread_merges_objects() {
    assert_eq!(run_js(r#"
const a = { x: 1, y: 2 };
const b = { y: 3, z: 4 };
const c = { ...a, ...b };
console.log(c.x);
console.log(c.y); // b wins
console.log(c.z);
"#), vec!["1", "3", "4"]);
}

#[test]
fn spread_creates_shallow_copy() {
    assert_eq!(run_js(r#"
const original = { a: { x: 1 } };
const copy = { ...original };
copy.a.x = 99;
console.log(original.a.x); // shared reference
"#), vec!["99"]);
}

#[test]
fn spread_does_not_copy_inherited() {
    assert_eq!(run_js(r#"
const proto = { inherited: true };
const src = Object.create(proto);
src.own = true;
const result = { ...src };
console.log(result.own);
console.log(result.inherited);
"#), vec!["true", "undefined"]);
}

#[test]
fn spread_does_not_copy_non_enumerable() {
    assert_eq!(run_js(r#"
const src = {};
Object.defineProperty(src, "hidden", { value: 1, enumerable: false });
const result = { ...src };
console.log(result.hidden);
"#), vec!["undefined"]);
}

#[test]
fn spread_copies_symbol_keys() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const src = { [sym]: 42, str: "ok" };
const result = { ...src };
console.log(result[sym]);
"#), vec!["42"]);
}

#[test]
fn spread_reads_getter_value() {
    assert_eq!(run_js(r#"
const src = { get x() { return 42; } };
const result = { ...src };
// result.x is plain data property, not a getter
const desc = Object.getOwnPropertyDescriptor(result, "x");
console.log(desc.value);
console.log(typeof desc.get);
"#), vec!["42", "undefined"]);
}

#[test]
fn spread_empty_is_noop() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
const result = { ...obj, ...{}, ...null }; // null/undefined are ignored
console.log(result.a);
"#), vec!["1"]);
}

#[test]
fn spread_overrides_with_explicit_property() {
    assert_eq!(run_js(r#"
const defaults = { color: "red", size: "M" };
const custom = { ...defaults, color: "blue" };
console.log(custom.color);
console.log(custom.size);
"#), vec!["blue", "M"]);
}

#[test]
fn spread_preserves_insertion_order() {
    assert_eq!(run_js(r#"
const result = { c: 3, ...{ a: 1, b: 2 } };
console.log(Object.keys(result).join(","));
"#), vec!["c,a,b"]);
}

#[test]
fn spread_of_string() {
    assert_eq!(run_js(r#"
const chars = { ..."abc" };
console.log(chars[0]);
console.log(chars[1]);
console.log(chars[2]);
"#), vec!["a", "b", "c"]);
}

#[test]
fn nested_spread_pattern() {
    assert_eq!(run_js(r#"
const state = { user: { name: "Alice", age: 30 }, count: 0 };
// Update nested immutably
const next = { ...state, user: { ...state.user, age: 31 }, count: state.count + 1 };
console.log(next.user.name);
console.log(next.user.age);
console.log(next.count);
console.log(state.user.age); // original unchanged
"#), vec!["Alice", "31", "1", "30"]);
}
