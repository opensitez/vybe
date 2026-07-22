use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Object.create()`, Prototypes & Property Descriptors
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_create_with_prototype_inheritance() {
    let src = r#"
const proto = { greeting: "Hello" };
const obj = Object.create(proto);
console.log(obj.greeting + "|isOwn=" + Object.hasOwn(obj, "greeting"));
"#;
    assert_eq!(run_js(src), vec!["Hello|isOwn=false"]);
}

#[test]
fn test_js_object_create_null_prototype_has_no_builtin_methods() {
    let src = r#"
const nullProtoObj = Object.create(null);
console.log(Object.getPrototypeOf(nullProtoObj) === null + "|hasToString=" + ("toString" in nullProtoObj));
"#;
    assert_eq!(run_js(src), vec!["true|hasToString=false"]);
}

#[test]
fn test_js_object_create_with_property_descriptors_map() {
    let src = r#"
const obj = Object.create(null, {
    x: { value: 10, writable: true, enumerable: true, configurable: true },
    y: { value: 20, writable: false, enumerable: false, configurable: false }
});
console.log(`${obj.x}:${obj.y}:${Object.keys(obj).join(",")}`);
"#;
    assert_eq!(run_js(src), vec!["10:20:x"]);
}

#[test]
fn test_js_object_create_getter_setter_descriptor_map() {
    let src = r#"
let store = 0;
const obj = Object.create(null, {
    val: {
        get() { return store; },
        set(v) { store = v * 2; },
        enumerable: true
    }
});
obj.val = 5;
console.log(obj.val);
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_object_create_invalid_prototype_throws_typeerror() {
    let src = r#"
try {
    Object.create(12345);
} catch (e) {
    console.log("Object.create Invalid Prototype TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Object.create Invalid Prototype TypeError"]
    );
}

#[test]
fn test_js_object_create_null_proto_dictionary_fast_lookup() {
    let src = r#"
const dict = Object.create(null);
dict.key1 = "v1";
dict["__proto__"] = "not_a_proto";
console.log(dict["__proto__"] + "|" + (Object.getPrototypeOf(dict) === null));
"#;
    assert_eq!(run_js(src), vec!["not_a_proto|true"]);
}

#[test]
fn test_js_object_get_prototype_of_and_set_prototype_of() {
    let src = r#"
const p1 = { a: 1 };
const p2 = { b: 2 };
const obj = Object.create(p1);
console.log(obj.a + "|b=" + obj.b);
Object.setPrototypeOf(obj, p2);
console.log("a=" + obj.a + "|b=" + obj.b);
"#;
    assert_eq!(run_js(src), vec!["1|b=undefined", "a=undefined|b=2"]);
}

#[test]
fn test_js_object_set_prototype_of_frozen_object_throws_typeerror() {
    let src = r#"
const obj = Object.freeze({});
try {
    Object.setPrototypeOf(obj, { newProto: true });
} catch (e) {
    console.log("SetPrototypeOf Frozen TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["SetPrototypeOf Frozen TypeError"]);
}

#[test]
fn test_js_object_set_prototype_of_prototype_cycle_throws_typeerror() {
    let src = r#"
const a = {};
const b = Object.create(a);
try {
    Object.setPrototypeOf(a, b); // Creating a prototype cycle is a TypeError!
} catch (e) {
    console.log("Prototype Cycle TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Prototype Cycle TypeError"]);
}

#[test]
fn test_js_object_is_prototype_of_chain_check() {
    let src = r#"
const grandParent = {};
const parent = Object.create(grandParent);
const child = Object.create(parent);

console.log(`${grandParent.isPrototypeOf(child)}:${parent.isPrototypeOf(child)}:${child.isPrototypeOf(parent)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:false"]);
}

#[test]
fn test_js_object_create_default_descriptor_boolean_flags_are_false() {
    let src = r#"
const obj = Object.create(null, {
    prop: { value: "defaultFlags" }
});
const desc = Object.getOwnPropertyDescriptor(obj, "prop");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:false"]); // Omitted descriptor boolean flags default to false!
}

#[test]
fn test_js_object_create_property_descriptor_must_be_object() {
    let src = r#"
try {
    Object.create(null, { prop: "not_an_object_descriptor" });
} catch (e) {
    console.log("Invalid Property Descriptor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Property Descriptor TypeError"]);
}

#[test]
fn test_js_object_create_conflicting_value_and_getter_throws_typeerror() {
    let src = r#"
try {
    Object.create(null, {
        invalid: { value: 10, get() { return 10; } }
    });
} catch (e) {
    console.log("Conflicting Descriptor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Conflicting Descriptor TypeError"]);
}

#[test]
fn test_js_object_create_conflicting_writable_and_setter_throws_typeerror() {
    let src = r#"
try {
    Object.create(null, {
        invalid: { writable: true, set(v) {} }
    });
} catch (e) {
    console.log("Writable Setter TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Writable Setter TypeError"]);
}

#[test]
fn test_js_object_create_symbol_property_descriptors() {
    let src = r#"
const sym = Symbol("symKey");
const obj = Object.create(null, {
    [sym]: { value: "symbolVal", enumerable: true }
});
console.log(obj[sym] + "|" + Object.getOwnPropertySymbols(obj).length);
"#;
    assert_eq!(run_js(src), vec!["symbolVal|1"]);
}

#[test]
fn test_js_object_create_multiple_levels_of_inheritance() {
    let src = r#"
const l1 = { level: 1 };
const l2 = Object.create(l1);
l2.level = 2;
const l3 = Object.create(l2);
console.log(`${l3.level}:${l2.level}:${l1.level}`);
"#;
    assert_eq!(run_js(src), vec!["2:2:1"]);
}

#[test]
fn test_js_object_create_with_empty_descriptors_map() {
    let src = r#"
const obj = Object.create({}, {});
console.log(Object.keys(obj).length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_object_create_undefined_descriptors_ignored() {
    let src = r#"
const obj = Object.create({ a: 1 }, undefined);
console.log(obj.a);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_object_set_prototype_of_same_prototype_is_noop() {
    let src = r#"
const proto = { x: 10 };
const obj = Object.create(proto);
const res = Object.setPrototypeOf(obj, proto);
console.log(res === obj);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_get_prototype_of_primitive_wrapper() {
    let src = r#"
console.log(Object.getPrototypeOf("str") === String.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
