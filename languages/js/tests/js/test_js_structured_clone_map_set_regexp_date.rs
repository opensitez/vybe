use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `structuredClone` with `Map`, `Set`, `RegExp` & `Date` Builtins
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_structured_clone_map_deep_copy() {
    let src = r#"
const map = new Map([["a", { val: 1 }]]);
const clone = structuredClone(map);
console.log((clone !== map) + "|" + (clone.get("a") !== map.get("a")) + "|" + (clone.get("a").val === 1));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_structured_clone_set_deep_copy() {
    let src = r#"
const obj = { id: 10 };
const set = new Set([obj]);
const clone = structuredClone(set);
const clonedObj = [...clone][0];
console.log((clone !== set) + "|" + (clonedObj !== obj) + "|" + (clonedObj.id === 10));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_structured_clone_date_object() {
    let src = r#"
const d = new Date("2026-07-22T12:00:00Z");
const clone = structuredClone(d);
console.log((clone !== d) + "|" + (clone instanceof Date) + "|" + (clone.getTime() === d.getTime()));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_structured_clone_regexp_object() {
    let src = r#"
const re = /test_\d+/giu;
const clone = structuredClone(re);
console.log((clone !== re) + "|" + (clone instanceof RegExp) + "|" + clone.source + "|" + clone.flags);
"#;
    assert_eq!(run_js(src), vec!["true|true|test_\\d+|giu"]);
}

#[test]
fn test_js_structured_clone_map_complex_keys() {
    let src = r#"
const key = { k: "objKey" };
const map = new Map([[key, "val"]]);
const clone = structuredClone(map);
const cloneKey = [...clone.keys()][0];
console.log((cloneKey !== key) + "|" + (clone.get(cloneKey) === "val"));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_invalid_date_nan() {
    let src = r#"
const d = new Date(NaN);
const clone = structuredClone(d);
console.log((clone instanceof Date) + "|" + Number.isNaN(clone.getTime()));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_regexp_lastindex_preserved() {
    let src = r#"
const re = /a/g;
re.lastIndex = 3;
const clone = structuredClone(re);
console.log(clone.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_structured_clone_weakmap_throws_datacloneerror() {
    let src = r#"
try {
    structuredClone(new WeakMap());
} catch (e) {
    console.log("DataCloneError WeakMap");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError WeakMap"]);
}

#[test]
fn test_js_structured_clone_weakset_throws_datacloneerror() {
    let src = r#"
try {
    structuredClone(new WeakSet());
} catch (e) {
    console.log("DataCloneError WeakSet");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError WeakSet"]);
}

#[test]
fn test_js_structured_clone_map_with_array_values() {
    let src = r#"
const map = new Map([["nums", [1, 2, 3]]]);
const clone = structuredClone(map);
clone.get("nums").push(4);
console.log(map.get("nums").length + "|" + clone.get("nums").length);
"#;
    assert_eq!(run_js(src), vec!["3|4"]);
}

#[test]
fn test_js_structured_clone_set_with_nested_maps() {
    let src = r#"
const innerMap = new Map([["x", 100]]);
const set = new Set([innerMap]);
const clone = structuredClone(set);
const clonedMap = [...clone][0];
console.log(clonedMap.get("x"));
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_structured_clone_regexp_named_groups_clone() {
    let src = r#"
const re = /(?<year>\d{4})-(?<month>\d{2})/g;
const clone = structuredClone(re);
const match = clone.exec("2026-07");
console.log(match.groups.year + "|" + match.groups.month);
"#;
    assert_eq!(run_js(src), vec!["2026|07"]);
}

#[test]
fn test_js_structured_clone_custom_properties_on_map_ignored() {
    let src = r#"
const map = new Map();
map.customProp = "customData";
const clone = structuredClone(map);
console.log(clone.size + "|hasCustomProp=" + Object.hasOwn(clone, "customProp"));
"#;
    assert_eq!(run_js(src), vec!["0|hasCustomProp=false"]);
}

#[test]
fn test_js_structured_clone_custom_properties_on_date_ignored() {
    let src = r#"
const d = new Date();
d.meta = 123;
const clone = structuredClone(d);
console.log(Object.hasOwn(clone, "meta"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_structured_clone_empty_map_and_set() {
    let src = r#"
const m = new Map();
const s = new Set();
const cm = structuredClone(m);
const cs = structuredClone(s);
console.log(cm.size + "|" + cs.size);
"#;
    assert_eq!(run_js(src), vec!["0|0"]);
}

#[test]
fn test_js_structured_clone_map_key_value_identity_preservation() {
    let src = r#"
const shared = { id: 1 };
const map = new Map([[shared, shared]]);
const clone = structuredClone(map);
const [cloneKey, cloneVal] = [...clone.entries()][0];
console.log((cloneKey !== shared) + "|" + (cloneKey === cloneVal));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_boolean_object_wrapper() {
    let src = r#"
const boolObj = new Boolean(true);
const clone = structuredClone(boolObj);
console.log((clone instanceof Boolean) + "|" + (clone.valueOf() === true));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_number_object_wrapper() {
    let src = r#"
const numObj = new Number(42);
const clone = structuredClone(numObj);
console.log((clone instanceof Number) + "|" + (clone.valueOf() === 42));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_string_object_wrapper() {
    let src = r#"
const strObj = new String("hello");
const clone = structuredClone(strObj);
console.log((clone instanceof String) + "|" + (clone.valueOf() === "hello"));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_regexp_unicode_sets_v_flag() {
    let src = r#"
const re = /[\p{Decimal_Number}]/v;
const clone = structuredClone(re);
console.log(clone.flags.includes("v"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
