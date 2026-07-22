use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Reflect API (`ownKeys`, `has`, `deleteProperty`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_reflect_ownkeys_string_and_symbol_keys() {
    let src = r#"
const sym = Symbol("s");
const obj = { a: 1, 0: "zero", [sym]: 2 };
const keys = Reflect.ownKeys(obj);
console.log(keys.map(k => String(k)).join(","));
"#;
    assert_eq!(run_js(src), vec!["0,a,Symbol(s)"]);
}

#[test]
fn test_js_reflect_has_property_in_operator_equivalent() {
    let src = r#"
const proto = { parentKey: 10 };
const obj = Object.create(proto);
obj.ownKey = 20;
console.log(Reflect.has(obj, "ownKey") + "|" + Reflect.has(obj, "parentKey") + "|" + Reflect.has(obj, "missing"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_reflect_deleteproperty_deletes_own_property() {
    let src = r#"
const obj = { a: 1, b: 2 };
const res = Reflect.deleteProperty(obj, "a");
console.log(res + "|hasA=" + ("a" in obj));
"#;
    assert_eq!(run_js(src), vec!["true|hasA=false"]);
}

#[test]
fn test_js_reflect_deleteproperty_non_configurable_returns_false() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 42, configurable: false });
const res = Reflect.deleteProperty(obj, "fixed");
console.log(res + "|" + obj.fixed);
"#;
    assert_eq!(run_js(src), vec!["false|42"]);
}

#[test]
fn test_js_reflect_ownkeys_includes_non_enumerable_properties() {
    let src = r#"
const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
const keys = Reflect.ownKeys(obj);
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["visible,hidden"]);
}

#[test]
fn test_js_reflect_has_symbol_property() {
    let src = r#"
const sym = Symbol("sym");
const obj = { [sym]: "Val" };
console.log(Reflect.has(obj, sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_reflect_deleteproperty_symbol_property() {
    let src = r#"
const sym = Symbol("sym");
const obj = { [sym]: "Val" };
const res = Reflect.deleteProperty(obj, sym);
console.log(res + "|hasSym=" + (sym in obj));
"#;
    assert_eq!(run_js(src), vec!["true|hasSym=false"]);
}

#[test]
fn test_js_reflect_ownkeys_ordering_canonical() {
    let src = r#"
const sym = Symbol("sym");
const obj = {
    "b": 1,
    "2": 2,
    "1": 1,
    [sym]: 3,
    "a": 0
};
const keys = Reflect.ownKeys(obj);
console.log(keys.map(k => String(k)).join(",")); // Numeric indices first (1, 2), string keys (b, a), then Symbols
"#;
    assert_eq!(run_js(src), vec!["1,2,b,a,Symbol(sym)"]);
}

#[test]
fn test_js_reflect_has_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.has("not_an_object", "length");
} catch (e) {
    console.log("Reflect.has Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect.has Non-Object TypeError"]);
}

#[test]
fn test_js_reflect_ownkeys_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.ownKeys(12345);
} catch (e) {
    console.log("Reflect.ownKeys Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect.ownKeys Non-Object TypeError"]);
}

#[test]
fn test_js_reflect_deleteproperty_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.deleteProperty(null, "key");
} catch (e) {
    console.log("Reflect.deleteProperty Non-Object TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Reflect.deleteProperty Non-Object TypeError"]
    );
}

#[test]
fn test_js_reflect_deleteproperty_missing_property_returns_true() {
    let src = r#"
const obj = {};
console.log(Reflect.deleteProperty(obj, "nonExistent"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_reflect_has_prototype_chain_traversal() {
    let src = r#"
class GrandParent {}
GrandParent.prototype.gpMethod = function() {};
class Parent extends GrandParent {}
class Child extends Parent {}

const c = new Child();
console.log(Reflect.has(c, "gpMethod"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_reflect_ownkeys_array_indices_and_length() {
    let src = r#"
const arr = [10, 20];
const keys = Reflect.ownKeys(arr);
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,length"]);
}

#[test]
fn test_js_reflect_deleteproperty_array_element() {
    let src = r#"
const arr = [10, 20, 30];
const res = Reflect.deleteProperty(arr, 1);
console.log(res + "|len=" + arr.length + "|hasIndex1=" + (1 in arr));
"#;
    assert_eq!(run_js(src), vec!["true|len=3|hasIndex1=false"]); // Creates a sparse array hole!
}

#[test]
fn test_js_reflect_ownkeys_function_object() {
    let src = r#"
function fn() {}
const keys = Reflect.ownKeys(fn);
console.log(keys.includes("length") + "|" + keys.includes("name") + "|" + keys.includes("prototype"));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_reflect_has_getter_setter_property() {
    let src = r#"
const obj = {
    get accessor() { return 1; }
};
console.log(Reflect.has(obj, "accessor"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_reflect_deleteproperty_getter_setter_property() {
    let src = r#"
const obj = {
    get accessor() { return 1; }
};
const res = Reflect.deleteProperty(obj, "accessor");
console.log(res + "|hasAccessor=" + ("accessor" in obj));
"#;
    assert_eq!(run_js(src), vec!["true|hasAccessor=false"]);
}

#[test]
fn test_js_reflect_ownkeys_empty_object() {
    let src = r#"
console.log(Reflect.ownKeys({}).length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_reflect_deleteproperty_sealed_object_returns_false() {
    let src = r#"
const obj = Object.seal({ prop: 10 });
const res = Reflect.deleteProperty(obj, "prop");
console.log(res + "|" + obj.prop);
"#;
    assert_eq!(run_js(src), vec!["false|10"]);
}
