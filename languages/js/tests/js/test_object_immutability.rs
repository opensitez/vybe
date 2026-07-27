/// Object.freeze deep patterns, immutability, nested objects
use super::helpers::run_js;

#[test]
fn freeze_prevents_property_modification() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.freeze({ x: 1, y: 2 });
obj.x = 99;
obj.z = 3;
console.log(obj.x);
console.log(obj.z);
"#
        ),
        vec!["1", "undefined"]
    );
}

#[test]
fn freeze_prevents_deletion() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.freeze({ x: 1 });
delete obj.x;
console.log(obj.x);
"#
        ),
        vec!["1"]
    );
}

#[test]
fn freeze_is_shallow_nested_mutable() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.freeze({ nested: { x: 1 } });
obj.nested.x = 99; // nested not frozen
console.log(obj.nested.x);
"#
        ),
        vec!["99"]
    );
}

#[test]
fn deep_freeze_pattern() {
    assert_eq!(
        run_js(
            r#"
function deepFreeze(obj) {
    Object.getOwnPropertyNames(obj).forEach(key => {
        const val = obj[key];
        if (typeof val === "object" && val !== null) deepFreeze(val);
    });
    return Object.freeze(obj);
}
const config = deepFreeze({ db: { host: "localhost", port: 5432 } });
config.db.port = 9999;
console.log(config.db.port);
"#
        ),
        vec!["5432"]
    );
}

#[test]
fn frozen_array_prevents_push() {
    assert_eq!(
        run_js(
            r#"
const arr = Object.freeze([1, 2, 3]);
try { arr.push(4); } catch {}
console.log(arr.length);
console.log(arr[3]);
"#
        ),
        vec!["3", "undefined"]
    );
}

#[test]
fn seal_allows_modify_prevents_add_delete() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.seal({ x: 1, y: 2 });
obj.x = 99; // modify ok
obj.z = 3;  // add — silently fails
delete obj.y; // delete — silently fails
console.log(obj.x);
console.log(obj.z);
console.log(obj.y);
"#
        ),
        vec!["99", "undefined", "2"]
    );
}

#[test]
fn seal_keeps_object_sealed_and_blocks_extension() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.seal({ x: 1 });
obj.x = 99;
obj.y = 2;
console.log(obj.x);
console.log(Object.prototype.hasOwnProperty.call(obj, "y"));
console.log(Object.isSealed(obj));
"#
        ),
        vec!["99", "false", "true"]
    );
}

#[test]
fn is_frozen_empty_non_extensible_is_frozen() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.freeze({});
console.log(Object.isFrozen(obj));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn object_create_null_no_prototype() {
    assert_eq!(
        run_js(
            r#"
const safe = Object.create(null);
safe.key = "value";
// No toString, no hasOwnProperty — truly property-free
console.log(typeof safe.toString);
console.log(safe.key);
"#
        ),
        vec!["undefined", "value"]
    );
}

#[test]
fn prevent_extensions_allows_modify() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.preventExtensions({ x: 1 });
obj.x = 99;  // existing properties can be modified
obj.y = 2;   // new properties silently fail
console.log(obj.x);
console.log(obj.y);
"#
        ),
        vec!["99", "undefined"]
    );
}

#[test]
fn prevent_extensions_freeze_status_checks() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
Object.preventExtensions(obj);
console.log(Object.isExtensible(obj));
console.log(Object.isSealed(obj));
console.log(Object.isFrozen(obj));
Object.freeze(obj);
console.log(Object.isExtensible(obj));
console.log(Object.isSealed(obj));
console.log(Object.isFrozen(obj));
"#
        ),
        vec!["false", "false", "false", "false", "true", "true"]
    );
}

#[test]
fn seal_makes_existing_properties_non_configurable() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.seal({ a: 1, b: 2 });
const d = Object.getOwnPropertyDescriptor(obj, "a");
console.log(d.configurable);
console.log(d.writable);
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn prevent_extensions_on_array_rejects_length_growth() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
Object.preventExtensions(arr);
let threw = false;
try {
    arr.push(4);
} catch {
    threw = true;
}
arr[4] = 5;
console.log(threw);
console.log(arr.length);
console.log(arr[3]);
console.log(arr[4]);
"#
        ),
        vec!["true", "3", "undefined", "undefined"]
    );
}
