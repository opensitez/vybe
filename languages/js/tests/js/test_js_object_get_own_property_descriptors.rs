use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object.getOwnPropertyDescriptors & Object.getOwnPropertyDescriptor
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_get_own_property_descriptor_normal_property() {
    let src = r#"
const obj = { a: 10 };
const desc = Object.getOwnPropertyDescriptor(obj, "a");
console.log(desc.value + "|" + desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["10|true|true|true"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_missing_property_returns_undefined() {
    let src = r#"
const obj = { a: 10 };
const desc = Object.getOwnPropertyDescriptor(obj, "missing");
console.log(desc === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_prototype_property_ignored() {
    let src = r#"
const proto = { protoProp: "parent" };
const obj = Object.create(proto);
obj.ownProp = "child";

console.log(Object.getOwnPropertyDescriptor(obj, "ownProp") !== undefined);
console.log(Object.getOwnPropertyDescriptor(obj, "protoProp") === undefined);
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_all_own_keys() {
    let src = r#"
const obj = { x: 1, y: 2 };
const descs = Object.getOwnPropertyDescriptors(obj);
console.log(Object.keys(descs).join(","));
console.log(descs.x.value + "|" + descs.y.value);
"#;
    assert_eq!(run_js(src), vec!["x,y", "1|2"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_includes_symbols() {
    let src = r#"
const sym = Symbol("id");
const obj = { [sym]: 100, name: "Item" };
const descs = Object.getOwnPropertyDescriptors(obj);
console.log(descs[sym].value + "|" + descs.name.value);
"#;
    assert_eq!(run_js(src), vec!["100|Item"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_clone_mixins() {
    let src = r#"
const source = {
    _count: 0,
    get count() { return this._count; },
    set count(v) { this._count = v; }
};
const clone = Object.create(Object.getPrototypeOf(source), Object.getOwnPropertyDescriptors(source));
clone.count = 50;
console.log(clone.count + "|" + source.count);
"#;
    assert_eq!(run_js(src), vec!["50|0"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_primitive_coercion() {
    let src = r#"
const descNumber = Object.getOwnPropertyDescriptor(100, "toString");
const descString = Object.getOwnPropertyDescriptor("abc", "length");
console.log(descNumber === undefined);
console.log(descString.value + "|" + descString.writable + "|" + descString.enumerable);
"#;
    assert_eq!(run_js(src), vec!["true", "3|false|false"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_accessor_properties() {
    let src = r#"
const obj = {
    get item() { return "value"; }
};
const desc = Object.getOwnPropertyDescriptor(obj, "item");
console.log(typeof desc.get + "|" + desc.set + "|" + desc.value);
"#;
    assert_eq!(run_js(src), vec!["function|undefined|undefined"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_empty_object() {
    let src = r#"
const descs = Object.getOwnPropertyDescriptors({});
console.log(Object.keys(descs).length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_non_enumerable_included() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "secret", { value: 42, enumerable: false });
const descs = Object.getOwnPropertyDescriptors(obj);
console.log(descs.secret.value + "|" + descs.secret.enumerable);
"#;
    assert_eq!(run_js(src), vec!["42|false"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_null_prototype() {
    let src = r#"
const obj = Object.create(null);
obj.key = "val";
const desc = Object.getOwnPropertyDescriptor(obj, "key");
console.log(desc.value);
"#;
    assert_eq!(run_js(src), vec!["val"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_symbol_property_descriptors() {
    let src = r#"
const sym = Symbol("sym");
const obj = {};
Object.defineProperty(obj, sym, { value: "SymbolVal", writable: false });
const descs = Object.getOwnPropertyDescriptors(obj);
console.log(descs[sym].writable);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_array_index() {
    let src = r#"
const arr = ["first", "second"];
const desc0 = Object.getOwnPropertyDescriptor(arr, 0);
const descLen = Object.getOwnPropertyDescriptor(arr, "length");
console.log(desc0.value + "|" + descLen.value + "|" + descLen.writable);
"#;
    assert_eq!(run_js(src), vec!["first|2|true"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_function_name_length() {
    let src = r#"
function testFn(a, b) {}
const descName = Object.getOwnPropertyDescriptor(testFn, "name");
const descLen = Object.getOwnPropertyDescriptor(testFn, "length");
console.log(descName.value + "|" + descLen.value + "|" + descName.writable);
"#;
    assert_eq!(run_js(src), vec!["testFn|2|false"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_shallow_copy() {
    let src = r#"
const orig = { a: 1 };
const descs = Object.getOwnPropertyDescriptors(orig);
descs.a.value = 99;
console.log(orig.a); // Original target property value unaltered
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_null_undefined_throws() {
    let src = r#"
try {
    Object.getOwnPropertyDescriptor(null, "key");
} catch (e) {
    console.log("Null Error");
}
try {
    Object.getOwnPropertyDescriptor(undefined, "key");
} catch (e) {
    console.log("Undefined Error");
}
"#;
    assert_eq!(run_js(src), vec!["Null Error", "Undefined Error"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_null_undefined_throws() {
    let src = r#"
try {
    Object.getOwnPropertyDescriptors(null);
} catch (e) {
    console.log("Descriptors Null Error");
}
"#;
    assert_eq!(run_js(src), vec!["Descriptors Null Error"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_frozen_object_descriptors() {
    let src = r#"
const obj = Object.freeze({ x: 10 });
const desc = Object.getOwnPropertyDescriptor(obj, "x");
console.log(desc.writable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false"]);
}

#[test]
fn test_js_object_get_own_property_descriptor_sealed_object_descriptors() {
    let src = r#"
const obj = Object.seal({ x: 10 });
const desc = Object.getOwnPropertyDescriptor(obj, "x");
console.log(desc.writable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_object_get_own_property_descriptors_class_prototype_methods() {
    let src = r#"
Class Foo {
    bar() {}
}
const descs = Object.getOwnPropertyDescriptors(Foo.prototype);
console.log(descs.bar.enumerable + "|" + descs.bar.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|true"]);
}
