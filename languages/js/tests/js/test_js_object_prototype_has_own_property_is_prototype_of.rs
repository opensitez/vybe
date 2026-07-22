use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object.prototype.hasOwnProperty, Object.hasOwn & isPrototypeOf
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_has_own_basic_property() {
    let src = r#"
const obj = { key: "value" };
console.log(Object.hasOwn(obj, "key"));
console.log(Object.hasOwn(obj, "missing"));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_has_own_property_prototype_inheritance() {
    let src = r#"
const parent = { parentKey: "parentVal" };
const child = Object.create(parent);
child.childKey = "childVal";

console.log(child.hasOwnProperty("childKey"));
console.log(child.hasOwnProperty("parentKey"));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_has_own_null_prototype_safety() {
    let src = r#"
const obj = Object.create(null);
obj.prop = 42;

// Object.hasOwn is safe on null prototype objects!
console.log(Object.hasOwn(obj, "prop"));

try {
    // Calling hasOwnProperty directly on object with null prototype throws TypeError!
    obj.hasOwnProperty("prop");
} catch (e) {
    console.log("Direct Call Failed");
}
"#;
    assert_eq!(run_js(src), vec!["true", "Direct Call Failed"]);
}

#[test]
fn test_js_object_has_own_overridden_has_own_property_safe() {
    let src = r#"
const obj = {
    hasOwnProperty: () => false,
    key: "value"
};
console.log(obj.hasOwnProperty("key")); // Overridden function returns false
console.log(Object.hasOwn(obj, "key")); // Object.hasOwn correctly inspects slot!
"#;
    assert_eq!(run_js(src), vec!["false", "true"]);
}

#[test]
fn test_js_object_has_own_symbol_keys() {
    let src = r#"
const sym = Symbol("id");
const obj = { [sym]: 100 };
console.log(Object.hasOwn(obj, sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_has_own_array_indices_and_length() {
    let src = r#"
const arr = [10, 20];
console.log(Object.hasOwn(arr, 0));
console.log(Object.hasOwn(arr, 1));
console.log(Object.hasOwn(arr, 2));
console.log(Object.hasOwn(arr, "length"));
"#;
    assert_eq!(run_js(src), vec!["true", "true", "false", "true"]);
}

#[test]
fn test_js_object_is_prototype_of_direct_parent() {
    let src = r#"
const proto = { a: 1 };
const child = Object.create(proto);
console.log(proto.isPrototypeOf(child));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_is_prototype_of_multilevel_chain() {
    let src = r#"
const grandParent = { g: 1 };
const parent = Object.create(grandParent);
const child = Object.create(parent);

console.log(grandParent.isPrototypeOf(child));
console.log(parent.isPrototypeOf(child));
console.log(Object.prototype.isPrototypeOf(child));
"#;
    assert_eq!(run_js(src), vec!["true", "true", "true"]);
}

#[test]
fn test_js_object_is_prototype_of_unrelated_objects() {
    let src = r#"
const a = {};
const b = {};
console.log(a.isPrototypeOf(b));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_object_is_prototype_of_class_instances() {
    let src = r#"
class Animal {}
class Dog extends Animal {}
const d = new Dog();

console.log(Dog.prototype.isPrototypeOf(d));
console.log(Animal.prototype.isPrototypeOf(d));
console.log(Object.prototype.isPrototypeOf(d));
"#;
    assert_eq!(run_js(src), vec!["true", "true", "true"]);
}

#[test]
fn test_js_object_has_own_non_enumerable_property() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "hidden", { value: 1, enumerable: false });
console.log(Object.hasOwn(obj, "hidden"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_has_own_primitive_coercion() {
    let src = r#"
console.log(Object.hasOwn("abc", 0));
console.log(Object.hasOwn("abc", "length"));
console.log(Object.hasOwn("abc", 3));
"#;
    assert_eq!(run_js(src), vec!["true", "true", "false"]);
}

#[test]
fn test_js_object_has_own_null_undefined_target_throws() {
    let src = r#"
try {
    Object.hasOwn(null, "key");
} catch (e) {
    console.log("HasOwn Null Error");
}
try {
    Object.hasOwn(undefined, "key");
} catch (e) {
    console.log("HasOwn Undefined Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["HasOwn Null Error", "HasOwn Undefined Error"]
    );
}

#[test]
fn test_js_object_is_prototype_of_primitive_targets_return_false() {
    let src = r#"
const proto = {};
console.log(proto.isPrototypeOf(42));
console.log(proto.isPrototypeOf("str"));
console.log(proto.isPrototypeOf(null));
"#;
    assert_eq!(run_js(src), vec!["false", "false", "false"]);
}

#[test]
fn test_js_object_property_is_enumerable_own_property() {
    let src = r#"
const obj = { visible: 1 };
Object.defineProperty(obj, "invisible", { value: 2, enumerable: false });

console.log(obj.propertyIsEnumerable("visible"));
console.log(obj.propertyIsEnumerable("invisible"));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_object_property_is_enumerable_inherited_property_returns_false() {
    let src = r#"
const parent = { parentProp: 1 };
const child = Object.create(parent);

console.log(child.propertyIsEnumerable("parentProp"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_object_property_is_enumerable_symbol_keys() {
    let src = r#"
const s1 = Symbol("s1");
const s2 = Symbol("s2");
const obj = { [s1]: 10 };
Object.defineProperty(obj, s2, { value: 20, enumerable: false });

console.log(obj.propertyIsEnumerable(s1));
console.log(obj.propertyIsEnumerable(s2));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
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
fn test_js_object_has_own_getter_setter_property() {
    let src = r#"
const obj = {
    get x() { return 10; }
};
console.log(Object.hasOwn(obj, "x"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_is_prototype_of_function_prototype() {
    let src = r#"
function Foo() {}
console.log(Function.prototype.isPrototypeOf(Foo));
console.log(Object.prototype.isPrototypeOf(Function.prototype));
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}
