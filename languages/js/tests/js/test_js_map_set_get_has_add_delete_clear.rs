use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Map & Set Core API Methods (get, set, add, delete, clear, size)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_map_set_get_has_delete_flow() {
    let src = r#"
const map = new Map();
map.set("key1", "val1");
map.set("key2", "val2");

console.log(map.get("key1") + "|" + map.has("key2") + "|size=" + map.size);
map.delete("key1");
console.log(map.has("key1") + "|size=" + map.size);
"#;
    assert_eq!(run_js(src), vec!["val1|true|size=2", "false|size=1"]);
}

#[test]
fn test_js_map_clear_resets_size() {
    let src = r#"
const map = new Map([["a", 1], ["b", 2]]);
map.clear();
console.log(map.size + "|" + map.has("a"));
"#;
    assert_eq!(run_js(src), vec!["0|false"]);
}

#[test]
fn test_js_map_object_and_function_keys() {
    let src = r#"
const objKey = { id: 1 };
const fnKey = function() {};
const map = new Map();
map.set(objKey, "ObjectVal");
map.set(fnKey, "FunctionVal");

console.log(map.get(objKey) + "|" + map.get(fnKey) + "|" + map.get({ id: 1 }));
"#;
    assert_eq!(run_js(src), vec!["ObjectVal|FunctionVal|undefined"]);
}

#[test]
fn test_js_map_nan_keys_same_value_zero() {
    let src = r#"
const map = new Map();
map.set(NaN, "NaN_Val");
console.log(map.get(NaN) + "|" + map.has(0 / 0));
"#;
    assert_eq!(run_js(src), vec!["NaN_Val|true"]);
}

#[test]
fn test_js_map_positive_and_negative_zero_keys_same_value_zero() {
    let src = r#"
const map = new Map();
map.set(+0, "Zero");
console.log(map.get(-0));
"#;
    assert_eq!(run_js(src), vec!["Zero"]);
}

#[test]
fn test_js_set_add_has_delete_flow() {
    let src = r#"
const set = new Set();
set.add(10);
set.add(20);
set.add(10); // Duplicate ignored

console.log(set.size + "|" + set.has(10));
set.delete(10);
console.log(set.size + "|" + set.has(10));
"#;
    assert_eq!(run_js(src), vec!["2|true", "1|false"]);
}

#[test]
fn test_js_set_clear_resets_size() {
    let src = r#"
const set = new Set([1, 2, 3]);
set.clear();
console.log(set.size);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_set_object_reference_uniqueness() {
    let src = r#"
const set = new Set();
const o1 = { a: 1 };
const o2 = { a: 1 };
set.add(o1);
set.add(o2);
set.add(o1); // Duplicate reference ignored

console.log(set.size);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_map_chainable_set_method() {
    let src = r#"
const map = new Map();
map.set("a", 1).set("b", 2).set("c", 3);
console.log(map.size);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_set_chainable_add_method() {
    let src = r#"
const set = new Set();
set.add("x").add("y").add("z");
console.log(set.size);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_map_constructor_with_iterable_tuples() {
    let src = r#"
const entries = [["k1", 10], ["k2", 20]];
const map = new Map(entries);
console.log(map.get("k1") + "|" + map.get("k2"));
"#;
    assert_eq!(run_js(src), vec!["10|20"]);
}

#[test]
fn test_js_set_constructor_with_iterable_array() {
    let src = r#"
const set = new Set(["apple", "banana", "apple"]);
console.log(set.size + "|" + set.has("banana"));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_map_delete_returns_boolean() {
    let src = r#"
const map = new Map([["exist", 1]]);
console.log(map.delete("exist") + "|" + map.delete("missing"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_set_delete_returns_boolean() {
    let src = r#"
const set = new Set([100]);
console.log(set.delete(100) + "|" + set.delete(200));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_map_symbol_keys() {
    let src = r#"
const s1 = Symbol("s1");
const s2 = Symbol("s2");
const map = new Map();
map.set(s1, "Val1");
map.set(s2, "Val2");

console.log(map.get(s1) + "|" + map.get(s2));
"#;
    assert_eq!(run_js(src), vec!["Val1|Val2"]);
}

#[test]
fn test_js_set_symbol_elements() {
    let src = r#"
const s1 = Symbol("s1");
const set = new Set([s1]);
console.log(set.has(s1) + "|" + set.has(Symbol("s1")));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_map_constructor_invalid_entry_throws_typeerror() {
    let src = r#"
try {
    new Map(["not_a_tuple"]);
} catch (e) {
    console.log("Map Entry Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Map Entry Non-Object TypeError"]);
}

#[test]
fn test_js_set_constructor_string_iterable() {
    let src = r#"
const set = new Set("hello");
console.log(set.size + "|" + Array.from(set).join(""));
"#;
    assert_eq!(run_js(src), vec!["4|helo"]);
}

#[test]
fn test_js_map_value_update_does_not_change_size() {
    let src = r#"
const map = new Map();
map.set("key", 100);
console.log(map.size);
map.set("key", 200); // Replaces existing value
console.log(map.get("key") + "|size=" + map.size);
"#;
    assert_eq!(run_js(src), vec!["1", "200|size=1"]);
}

#[test]
fn test_js_map_null_and_undefined_keys() {
    let src = r#"
const map = new Map();
map.set(null, "NullVal");
map.set(undefined, "UndefVal");

console.log(map.get(null) + "|" + map.get(undefined));
"#;
    assert_eq!(run_js(src), vec!["NullVal|UndefVal"]);
}
