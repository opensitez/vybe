/// Object.keys/values/entries ordering rules — integer indices first, then insertion order

use super::helpers::run_js;

#[test]
fn keys_integer_indices_sorted_first() {
    assert_eq!(run_js(r#"
const obj = { b: 2, 0: "zero", a: 1, 2: "two", 1: "one" };
const keys = Object.keys(obj);
console.log(keys.join(","));
"#), vec!["0,1,2,b,a"]);
}

#[test]
fn keys_string_properties_insertion_order() {
    assert_eq!(run_js(r#"
const obj = {};
obj.c = 3;
obj.a = 1;
obj.b = 2;
console.log(Object.keys(obj).join(","));
"#), vec!["c,a,b"]);
}

#[test]
fn keys_skips_symbols() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const obj = { a: 1, [sym]: 2, b: 3 };
console.log(Object.keys(obj).join(","));
"#), vec!["a,b"]);
}

#[test]
fn values_follows_keys_order() {
    assert_eq!(run_js(r#"
const obj = { b: 20, a: 10, c: 30 };
console.log(Object.values(obj).join(","));
"#), vec!["20,10,30"]);
}

#[test]
fn entries_returns_key_value_pairs() {
    assert_eq!(run_js(r#"
const obj = { x: 1, y: 2 };
const entries = Object.entries(obj);
console.log(entries.map(([k, v]) => k + "=" + v).join(","));
"#), vec!["x=1,y=2"]);
}

#[test]
fn get_own_property_names_includes_non_enumerable() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
const names = Object.getOwnPropertyNames(obj);
console.log(names.includes("a"));
console.log(names.includes("hidden"));
"#), vec!["true", "true"]);
}

#[test]
fn reflect_own_keys_all_types_ordered() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const obj = { 1: "b", sym: "s", 0: "a" };
obj[sym] = "sym";
const keys = Reflect.ownKeys(obj);
// integer indices first (0, 1), then string ("sym"), then Symbol
console.log(keys[0]); // "0"
console.log(keys[1]); // "1"
console.log(keys[2]); // "sym" (string)
console.log(typeof keys[3]); // symbol
"#), vec!["0", "1", "sym", "symbol"]);
}

#[test]
fn for_in_vs_object_keys_same_order() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2, c: 3 };
const forInKeys = [];
for (const k in obj) forInKeys.push(k);
const objectKeys = Object.keys(obj);
console.log(forInKeys.join(",") === objectKeys.join(","));
"#), vec!["true"]);
}

#[test]
fn json_stringify_key_order() {
    assert_eq!(run_js(r#"
// JSON.stringify follows own enumerable insertion order (non-integer)
const obj = { b: 2, a: 1, c: 3 };
const json = JSON.stringify(obj);
console.log(json);
"#), vec!["{\"b\":2,\"a\":1,\"c\":3}"]);
}

#[test]
fn negative_integer_string_not_sorted_as_index() {
    assert_eq!(run_js(r#"
const obj = { "-1": "neg", 0: "zero", a: "a" };
const keys = Object.keys(obj);
// -1 is not an array index, treated as string property
console.log(keys[0]); // 0 (integer index first)
console.log(keys.includes("-1"));
"#), vec!["0", "true"]);
}
