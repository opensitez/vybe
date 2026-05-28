/// Object.fromEntries patterns, transformation pipelines

use super::helpers::run_js;

#[test]
fn from_entries_from_array() {
    assert_eq!(run_js(r#"
const entries = [["a", 1], ["b", 2], ["c", 3]];
const obj = Object.fromEntries(entries);
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#), vec!["1", "2", "3"]);
}

#[test]
fn from_entries_from_map() {
    assert_eq!(run_js(r#"
const map = new Map([["x", 10], ["y", 20]]);
const obj = Object.fromEntries(map);
console.log(obj.x);
console.log(obj.y);
"#), vec!["10", "20"]);
}

#[test]
fn entries_then_from_entries_roundtrip() {
    assert_eq!(run_js(r#"
const original = { a: 1, b: 2, c: 3 };
const clone = Object.fromEntries(Object.entries(original));
console.log(clone.a);
console.log(clone.b);
console.log(clone.c);
"#), vec!["1", "2", "3"]);
}

#[test]
fn transform_object_values() {
    assert_eq!(run_js(r#"
const prices = { apple: 1.5, banana: 0.75, cherry: 2.0 };
const doubled = Object.fromEntries(
    Object.entries(prices).map(([k, v]) => [k, v * 2])
);
console.log(doubled.apple);
console.log(doubled.banana);
"#), vec!["3", "1.5"]);
}

#[test]
fn filter_object_by_value() {
    assert_eq!(run_js(r#"
const scores = { alice: 95, bob: 72, charlie: 88, dave: 61 };
const passing = Object.fromEntries(
    Object.entries(scores).filter(([, score]) => score >= 80)
);
console.log(Object.keys(passing).sort().join(","));
"#), vec!["alice,charlie"]);
}

#[test]
fn invert_object() {
    assert_eq!(run_js(r#"
const obj = { a: "1", b: "2", c: "3" };
const inverted = Object.fromEntries(
    Object.entries(obj).map(([k, v]) => [v, k])
);
console.log(inverted["1"]);
console.log(inverted["2"]);
"#), vec!["a", "b"]);
}

#[test]
fn from_entries_from_url_search_params() {
    assert_eq!(run_js(r#"
const params = new URLSearchParams("a=1&b=2&c=3");
const obj = Object.fromEntries(params);
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#), vec!["1", "2", "3"]);
}

#[test]
fn from_entries_with_generator() {
    assert_eq!(run_js(r#"
function* makeEntries() {
    yield ["x", 1];
    yield ["y", 2];
    yield ["z", 3];
}
const obj = Object.fromEntries(makeEntries());
console.log(obj.x);
console.log(obj.z);
"#), vec!["1", "3"]);
}

#[test]
fn pick_keys_from_object() {
    assert_eq!(run_js(r#"
function pick(obj, ...keys) {
    return Object.fromEntries(
        keys.filter(k => k in obj).map(k => [k, obj[k]])
    );
}
const user = { id: 1, name: "Alice", password: "secret", email: "alice@example.com" };
const safe = pick(user, "id", "name", "email");
console.log(Object.keys(safe).sort().join(","));
console.log(safe.name);
console.log("password" in safe);
"#), vec!["email,id,name", "Alice", "false"]);
}

#[test]
fn rename_keys() {
    assert_eq!(run_js(r#"
const keyMap = { firstName: "first_name", lastName: "last_name" };
const user = { firstName: "Alice", lastName: "Smith" };
const renamed = Object.fromEntries(
    Object.entries(user).map(([k, v]) => [keyMap[k] || k, v])
);
console.log(renamed.first_name);
console.log(renamed.last_name);
"#), vec!["Alice", "Smith"]);
}
