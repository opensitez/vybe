use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object Immutability (preventExtensions, seal, freeze & isExtensible/isSealed/isFrozen)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_prevent_extensions_prevents_new_properties() {
    let src = r#"
const obj = { existing: 1 };
Object.preventExtensions(obj);
console.log(Object.isExtensible(obj));
try {
    "use strict";
    obj.newProp = 2;
} catch (e) {
    console.log("PreventExtensions Error");
}
console.log(obj.newProp);
"#;
    assert_eq!(
        run_js(src),
        vec!["false", "PreventExtensions Error", "undefined"]
    );
}

#[test]
fn test_js_object_prevent_extensions_allows_modifying_existing() {
    let src = r#"
const obj = { val: 10 };
Object.preventExtensions(obj);
obj.val = 20;
console.log(obj.val);
delete obj.val;
console.log(obj.val);
"#;
    assert_eq!(run_js(src), vec!["20", "undefined"]);
}

#[test]
fn test_js_object_seal_prevents_add_delete_and_config() {
    let src = r#"
const obj = { a: 1 };
Object.seal(obj);
console.log(Object.isSealed(obj));
console.log(Object.isExtensible(obj));

obj.a = 100; // Modifying existing writable property allowed
console.log(obj.a);

try {
    "use strict";
    delete obj.a; // Deleting property throws in strict mode
} catch (e) {
    console.log("Seal Delete Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["true", "false", "100", "Seal Delete Error"]
    );
}

#[test]
fn test_js_object_freeze_prevents_all_mutations() {
    let src = r#"
const obj = { x: 10, y: "hello" };
Object.freeze(obj);
console.log(Object.isFrozen(obj));
console.log(Object.isSealed(obj));
console.log(Object.isExtensible(obj));

try {
    "use strict";
    obj.x = 20;
} catch (e) {
    console.log("Freeze Mutation Error");
}
console.log(obj.x);
"#;
    assert_eq!(
        run_js(src),
        vec!["true", "true", "false", "Freeze Mutation Error", "10"]
    );
}

#[test]
fn test_js_object_freeze_shallow_only() {
    let src = r#"
const obj = {
    nested: { val: 5 }
};
Object.freeze(obj);
obj.nested.val = 50; // Nested object is NOT frozen!
console.log(obj.nested.val);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_object_is_extensible_primitives_es6() {
    let src = r#"
console.log(Object.isExtensible(42));
console.log(Object.isExtensible("str"));
console.log(Object.isExtensible(null));
"#;
    assert_eq!(run_js(src), vec!["false", "false", "false"]);
}

#[test]
fn test_js_object_is_sealed_primitives_es6() {
    let src = r#"
console.log(Object.isSealed(42));
console.log(Object.isSealed("hello"));
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_object_is_frozen_primitives_es6() {
    let src = r#"
console.log(Object.isFrozen(42));
console.log(Object.isFrozen("world"));
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_object_freeze_array_mutations_throw() {
    let src = r#"
const arr = [1, 2, 3];
Object.freeze(arr);
console.log(Object.isFrozen(arr));
try {
    "use strict";
    arr[0] = 99;
} catch (e) {
    console.log("Frozen Array Element Error");
}
try {
    arr.push(4);
} catch (e) {
    console.log("Frozen Array Push Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec![
            "true",
            "Frozen Array Element Error",
            "Frozen Array Push Error"
        ]
    );
}

#[test]
fn test_js_object_prevent_extensions_on_empty_object_is_sealed_and_frozen() {
    let src = r#"
const obj = {};
Object.preventExtensions(obj);
console.log(Object.isExtensible(obj));
console.log(Object.isSealed(obj));
console.log(Object.isFrozen(obj));
"#;
    assert_eq!(run_js(src), vec!["false", "true", "true"]);
}

#[test]
fn test_js_object_seal_on_object_with_non_writable_is_frozen() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "prop", {
    value: 10,
    writable: false,
    configurable: true
});
Object.seal(obj);
console.log(Object.isFrozen(obj));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_freeze_getter_setter_preserves_accessor() {
    let src = r#"
const obj = {
    _count: 0,
    get count() { return this._count; },
    set count(v) { this._count = v; }
};
Object.freeze(obj);
// Calling setter mutates backing field because accessor properties are frozen without changing getter/setter pointers
obj.count = 10;
console.log(obj.count);
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_object_prevent_extensions_prototype_mutation() {
    let src = r#"
const proto = { parentVal: 100 };
const obj = Object.create(proto);
Object.preventExtensions(obj);

try {
    // Modern ES6 Object.setPrototypeOf on non-extensible throws TypeError!
    Object.setPrototypeOf(obj, { newProto: 200 });
} catch (e) {
    console.log("SetPrototypeOf Failed");
}
"#;
    assert_eq!(run_js(src), vec!["SetPrototypeOf Failed"]);
}

#[test]
fn test_js_object_freeze_returns_same_object_reference() {
    let src = r#"
const obj = { a: 1 };
const frozen = Object.freeze(obj);
console.log(obj === frozen);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_seal_returns_same_object_reference() {
    let src = r#"
const obj = { a: 1 };
const sealed = Object.seal(obj);
console.log(obj === sealed);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_prevent_extensions_returns_same_object_reference() {
    let src = r#"
const obj = { a: 1 };
const res = Object.preventExtensions(obj);
console.log(obj === res);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_freeze_symbol_properties() {
    let src = r#"
const sym = Symbol("id");
const obj = { [sym]: "Original" };
Object.freeze(obj);

try {
    "use strict";
    obj[sym] = "Mutated";
} catch (e) {
    console.log("Frozen Symbol Property Error");
}
console.log(obj[sym]);
"#;
    assert_eq!(
        run_js(src),
        vec!["Frozen Symbol Property Error", "Original"]
    );
}

#[test]
fn test_js_object_deep_freeze_implementation() {
    let src = r#"
function deepFreeze(obj) {
    Object.keys(obj).forEach(key => {
        if (typeof obj[key] === "object" && obj[key] !== null) deepFreeze(obj[key]);
    });
    return Object.freeze(obj);
}
const complex = { inner: { value: 42 } };
deepFreeze(complex);
console.log(Object.isFrozen(complex.inner));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_prevent_extensions_array_push_throws() {
    let src = r#"
const arr = [1, 2];
Object.preventExtensions(arr);
try {
    arr.push(3);
} catch (e) {
    console.log("Push Non-Extensible Array Error");
}
"#;
    assert_eq!(run_js(src), vec!["Push Non-Extensible Array Error"]);
}

#[test]
fn test_js_object_prevent_extensions_defines_property_throws() {
    let src = r#"
const obj = { x: 1 };
Object.preventExtensions(obj);
try {
    Object.defineProperty(obj, "y", { value: 2 });
} catch (e) {
    console.log("DefineProperty Extension Error");
}
"#;
    assert_eq!(run_js(src), vec!["DefineProperty Extension Error"]);
}
