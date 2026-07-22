use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object.defineProperty, Getter/Setter Descriptors & Attributes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_define_property_value_writable_false() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "prop", {
    value: 42,
    writable: false,
    configurable: true,
    enumerable: true
});
console.log(obj.prop);
try {
    "use strict";
    obj.prop = 99;
} catch (e) {
    console.log("Error: " + e.name);
}
console.log(obj.prop);
"#;
    assert_eq!(run_js(src), vec!["42", "Error: TypeError", "42"]);
}

#[test]
fn test_js_object_define_property_enumerable_false_in_keys() {
    let src = r#"
const obj = { a: 1 };
Object.defineProperty(obj, "hidden", {
    value: 2,
    enumerable: false
});
console.log(Object.keys(obj).join(","));
console.log(obj.hidden);
"#;
    assert_eq!(run_js(src), vec!["a", "2"]);
}

#[test]
fn test_js_object_define_property_configurable_false_redefine_throws() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", {
    value: 100,
    configurable: false
});
try {
    Object.defineProperty(obj, "fixed", { configurable: true });
} catch (e) {
    console.log("Cannot reconfigure: " + e.name);
}
"#;
    assert_eq!(run_js(src), vec!["Cannot reconfigure: TypeError"]);
}

#[test]
fn test_js_object_define_property_getter_setter_accessors() {
    let src = r#"
const obj = { _val: 0 };
Object.defineProperty(obj, "val", {
    get() { return this._val * 2; },
    set(v) { this._val = v + 5; },
    enumerable: true,
    configurable: true
});
obj.val = 10;
console.log(obj._val);
console.log(obj.val);
"#;
    assert_eq!(run_js(src), vec!["15", "30"]);
}

#[test]
fn test_js_object_define_property_getter_without_setter_strict_throw() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "readOnly", {
    get() { return "Constant"; },
    configurable: true
});
console.log(obj.readOnly);
try {
    "use strict";
    obj.readOnly = "NewVal";
} catch (e) {
    console.log("TypeError Caught");
}
"#;
    assert_eq!(run_js(src), vec!["Constant", "TypeError Caught"]);
}

#[test]
fn test_js_object_define_properties_multiple() {
    let src = r#"
const obj = {};
Object.defineProperties(obj, {
    x: { value: 10, writable: true, enumerable: true },
    y: { get() { return this.x * 2; }, enumerable: true }
});
console.log(obj.x + "," + obj.y);
obj.x = 20;
console.log(obj.y);
"#;
    assert_eq!(run_js(src), vec!["10,20", "40"]);
}

#[test]
fn test_js_object_define_property_defaults_all_false() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "raw", { value: 50 });
const desc = Object.getOwnPropertyDescriptor(obj, "raw");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_object_define_property_symbol_key() {
    let src = r#"
const sym = Symbol("privateKey");
const obj = {};
Object.defineProperty(obj, sym, {
    value: "Secret",
    writable: true,
    enumerable: true
});
console.log(obj[sym]);
console.log(Object.getOwnPropertySymbols(obj).length);
"#;
    assert_eq!(run_js(src), vec!["Secret", "1"]);
}

#[test]
fn test_js_object_define_property_accessor_value_conflict_throws() {
    let src = r#"
const obj = {};
try {
    Object.defineProperty(obj, "bad", {
        value: 10,
        get() { return 10; }
    });
} catch (e) {
    console.log("Conflict: " + e.name);
}
"#;
    assert_eq!(run_js(src), vec!["Conflict: TypeError"]);
}

#[test]
fn test_js_object_define_property_accessor_writable_conflict_throws() {
    let src = r#"
const obj = {};
try {
    Object.defineProperty(obj, "bad", {
        writable: true,
        set(v) {}
    });
} catch (e) {
    console.log("Conflict: " + e.name);
}
"#;
    assert_eq!(run_js(src), vec!["Conflict: TypeError"]);
}

#[test]
fn test_js_object_define_property_inheritance_accessor_receiver_this() {
    let src = r#"
const parent = {};
Object.defineProperty(parent, "name", {
    get() { return this._name || "Default"; },
    set(v) { this._name = v.toUpperCase(); },
    configurable: true
});
const child = Object.create(parent);
child.name = "alice";
console.log(child.name + "|" + parent.name);
"#;
    assert_eq!(run_js(src), vec!["ALICE|Default"]);
}

#[test]
fn test_js_object_define_property_delete_configurable_true() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "temp", {
    value: 100,
    configurable: true
});
console.log(delete obj.temp);
console.log(obj.temp);
"#;
    assert_eq!(run_js(src), vec!["true", "undefined"]);
}

#[test]
fn test_js_object_define_property_delete_configurable_false() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "perm", {
    value: 100,
    configurable: false
});
console.log(delete obj.perm);
console.log(obj.perm);
"#;
    assert_eq!(run_js(src), vec!["false", "100"]);
}

#[test]
fn test_js_object_define_property_value_modification_writable_true() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "counter", {
    value: 1,
    writable: true
});
obj.counter += 5;
console.log(obj.counter);
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_object_define_property_on_primitive_throws_typeerror() {
    let src = r#"
try {
    Object.defineProperty(42, "prop", { value: 1 });
} catch (e) {
    console.log("TypeError on Primitive");
}
"#;
    assert_eq!(run_js(src), vec!["TypeError on Primitive"]);
}

#[test]
fn test_js_object_define_property_redefine_same_value_non_writable() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "v", {
    value: 99,
    writable: false,
    configurable: false
});
// Redefining with identical value & writable: false succeeds!
Object.defineProperty(obj, "v", { value: 99 });
console.log(obj.v);
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_object_define_property_redefine_different_value_non_writable_throws() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "v", {
    value: 99,
    writable: false,
    configurable: false
});
try {
    Object.defineProperty(obj, "v", { value: 100 });
} catch (e) {
    console.log("Redefine Failed");
}
"#;
    assert_eq!(run_js(src), vec!["Redefine Failed"]);
}

#[test]
fn test_js_object_define_property_writable_true_to_false_transition() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "v", {
    value: 10,
    writable: true,
    configurable: false
});
// Transitioning writable true -> false is allowed even if configurable: false!
Object.defineProperty(obj, "v", { writable: false });
console.log(Object.getOwnPropertyDescriptor(obj, "v").writable);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_object_define_property_array_length_truncation() {
    let src = r#"
const arr = [1, 2, 3, 4, 5];
Object.defineProperty(arr, "length", { value: 2 });
console.log(arr.length + "|" + arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["2|1,2"]);
}

#[test]
fn test_js_object_define_property_array_length_non_writable() {
    let src = r#"
const arr = [10, 20];
Object.defineProperty(arr, "length", { writable: false });
try {
    arr.push(30);
} catch (e) {
    console.log("Length Non-Writable Error");
}
"#;
    assert_eq!(run_js(src), vec!["Length Non-Writable Error"]);
}
