use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Object.hasOwn()` vs `Object.prototype.hasOwnProperty()`
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_has_own_basic_own_property() {
    let src = r#"
const obj = { ownProp: 42 };
console.log(`${Object.hasOwn(obj, "ownProp")}:${Object.hasOwn(obj, "missing")}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_object_has_own_ignores_prototype_chain() {
    let src = r#"
const proto = { inheritedProp: "hello" };
const obj = Object.create(proto);
obj.ownProp = "world";

console.log(`${Object.hasOwn(obj, "ownProp")}:${Object.hasOwn(obj, "inheritedProp")}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_object_has_own_null_prototype_objects() {
    let src = r#"
const nullProtoObj = Object.create(null);
nullProtoObj.key = 100;
console.log(Object.hasOwn(nullProtoObj, "key")); // Object.hasOwn works safely on null-prototype objects!
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_has_own_property_fails_on_null_prototype_objects() {
    let src = r#"
const nullProtoObj = Object.create(null);
nullProtoObj.key = 100;
try {
    nullProtoObj.hasOwnProperty("key");
} catch (e) {
    console.log("hasOwnProperty Null Prototype TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["hasOwnProperty Null Prototype TypeError"]);
}

#[test]
fn test_js_object_has_own_overridden_has_own_property() {
    let src = r#"
const obj = {
    hasOwnProperty() { return false; }, // Overridden hasOwnProperty method!
    validKey: 99
};
console.log(Object.hasOwn(obj, "validKey") + "|" + obj.hasOwnProperty("validKey"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_object_has_own_symbol_properties() {
    let src = r#"
const sym = Symbol("id");
const obj = { [sym]: "symbolValue" };
console.log(Object.hasOwn(obj, sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_has_own_array_indices_and_length() {
    let src = r#"
const arr = [10, 20];
console.log(`${Object.hasOwn(arr, 0)}:${Object.hasOwn(arr, 1)}:${Object.hasOwn(arr, 2)}:${Object.hasOwn(arr, "length")}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false:true"]);
}

#[test]
fn test_js_object_has_own_sparse_array_holes() {
    let src = r#"
const sparse = [10, , 30];
console.log(`${Object.hasOwn(sparse, 0)}:${Object.hasOwn(sparse, 1)}:${Object.hasOwn(sparse, 2)}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true"]); // Sparse hole index 1 returns false!
}

#[test]
fn test_js_object_has_own_non_enumerable_properties() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "hidden", { value: "secret", enumerable: false });
console.log(Object.hasOwn(obj, "hidden"));
"#;
    assert_eq!(run_js(src), vec!["true"]); // Returns true even for non-enumerable own properties!
}

#[test]
fn test_js_object_has_own_getter_setter_properties() {
    let src = r#"
const obj = {
    get accessor() { return 1; }
};
console.log(Object.hasOwn(obj, "accessor"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_has_own_coerces_primitive_target_to_object() {
    let src = r#"
console.log(`${Object.hasOwn("hello", 0)}:${Object.hasOwn("hello", "length")}:${Object.hasOwn(123, "toString")}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_object_has_own_null_or_undefined_target_throws_typeerror() {
    let src = r#"
try {
    Object.hasOwn(null, "prop");
} catch (e) {
    console.log("hasOwn Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["hasOwn Null TypeError"]);
}

#[test]
fn test_js_object_has_own_coerces_property_key_to_string_or_symbol() {
    let src = r#"
const obj = { 42: "answer" };
console.log(Object.hasOwn(obj, 42) + "|" + Object.hasOwn(obj, "42"));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_object_has_own_function_properties() {
    let src = r#"
function fn() {}
console.log(`${Object.hasOwn(fn, "name")}:${Object.hasOwn(fn, "length")}:${Object.hasOwn(fn, "prototype")}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_object_has_own_class_static_and_instance_fields() {
    let src = r#"
class Widget {
    static staticProp = 1;
    instanceProp = 2;
}
const w = new Widget();
console.log(`${Object.hasOwn(Widget, "staticProp")}:${Object.hasOwn(w, "instanceProp")}:${Object.hasOwn(w, "staticProp")}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_object_has_own_deleted_property_returns_false() {
    let src = r#"
const obj = { a: 1 };
delete obj.a;
console.log(Object.hasOwn(obj, "a"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_object_has_own_property_descriptor_details() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Object, "hasOwn");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Object.hasOwn.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:2"]);
}

#[test]
fn test_js_object_has_own_name_property() {
    let src = r#"
console.log(Object.hasOwn.name);
"#;
    assert_eq!(run_js(src), vec!["hasOwn"]);
}

#[test]
fn test_js_object_has_own_typed_array_indices() {
    let src = r#"
const u8 = new Uint8Array([5, 10]);
console.log(`${Object.hasOwn(u8, 0)}:${Object.hasOwn(u8, 1)}:${Object.hasOwn(u8, 2)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_object_has_own_map_and_set_properties() {
    let src = r#"
const set = new Set();
console.log(`${Object.hasOwn(set, "size")}:${Object.hasOwn(Set.prototype, "size")}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]); // 'size' is an accessor property on Set.prototype, not own property of set instance!
}
