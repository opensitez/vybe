/// Property descriptor patterns — defineProperty, accessor vs data, writable, configurable
use super::helpers::run_js;

#[test]
fn define_non_writable_property() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "x", { value: 42, writable: false, configurable: true, enumerable: true });
obj.x = 99; // silently fails in sloppy mode
console.log(obj.x);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn define_non_enumerable_hidden_from_keys() {
    assert_eq!(
        run_js(
            r#"
const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false, configurable: true, writable: true });
console.log(Object.keys(obj).join(","));
console.log(obj.hidden);
"#
        ),
        vec!["visible", "2"]
    );
}

#[test]
fn define_accessor_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { _x: 0 };
Object.defineProperty(obj, "x", {
    get() { return this._x; },
    set(v) { this._x = v * 2; },
    configurable: true, enumerable: true
});
obj.x = 5;
console.log(obj.x);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn redefine_configurable_property() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "p", { value: 1, configurable: true, writable: true });
obj.p = 3;
console.log(obj.p);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn redefine_non_configurable_throws() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "p", { value: 1, configurable: false });
let threw = false;
try {
    Object.defineProperty(obj, "p", { value: 2 });
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn get_own_property_descriptor_returns_all_attributes() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
const desc = Object.getOwnPropertyDescriptor(obj, "x");
console.log(desc.value);
console.log(desc.writable);
console.log(desc.enumerable);
console.log(desc.configurable);
"#
        ),
        vec!["1", "true", "true", "true"]
    );
}

#[test]
fn accessor_descriptor_has_no_value() {
    assert_eq!(
        run_js(
            r#"
const obj = { get x() { return 1; } };
const desc = Object.getOwnPropertyDescriptor(obj, "x");
console.log("value" in desc);
console.log(typeof desc.get);
console.log(typeof desc.set);
"#
        ),
        vec!["false", "function", "undefined"]
    );
}

#[test]
fn define_properties_batch() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperties(obj, {
    a: { value: 1, enumerable: true, configurable: true, writable: true },
    b: { value: 2, enumerable: true, configurable: true, writable: true },
});
console.log(obj.a + obj.b);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn seal_prevents_new_and_makes_non_configurable() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
Object.seal(obj);
obj.y = 2; // silently fails
console.log(obj.y);
obj.x = 99; // still writable
console.log(obj.x);
const keyCount = Object.keys(obj).length;
obj.z = 3; // try adding another property
console.log(Object.keys(obj).length === keyCount); // sealed — no new keys
"#
        ),
        vec!["undefined", "99", "true"]
    );
}

#[test]
fn freeze_prevents_write_and_configure() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
Object.freeze(obj);
obj.x = 99;
obj.y = 2;
console.log(obj.x);
console.log(obj.y);
console.log(Object.isFrozen(obj));
"#
        ),
        vec!["1", "undefined", "true"]
    );
}

#[test]
fn is_sealed_and_is_frozen() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
console.log(Object.isSealed(obj));  // false
console.log(Object.isFrozen(obj));  // false (empty non-extensible would be frozen/sealed)
Object.seal(obj);
console.log(Object.isSealed(obj));  // true
"#
        ),
        vec!["false", "false", "true"]
    );
}

#[test]
fn prevent_extensions_blocks_new_props() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1 };
Object.preventExtensions(obj);
obj.b = 2; // silently fails
console.log(obj.b);
console.log(Object.isExtensible(obj));
"#
        ),
        vec!["undefined", "false"]
    );
}
