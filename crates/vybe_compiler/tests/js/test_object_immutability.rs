/// Object.freeze deep patterns, immutability, nested objects

use super::helpers::run_js;

#[test]
fn freeze_prevents_property_modification() {
    assert_eq!(run_js(r#"
const obj = Object.freeze({ x: 1, y: 2 });
obj.x = 99;
obj.z = 3;
console.log(obj.x);
console.log(obj.z);
"#), vec!["1", "undefined"]);
}

#[test]
fn freeze_prevents_deletion() {
    assert_eq!(run_js(r#"
const obj = Object.freeze({ x: 1 });
delete obj.x;
console.log(obj.x);
"#), vec!["1"]);
}

#[test]
fn freeze_is_shallow_nested_mutable() {
    assert_eq!(run_js(r#"
const obj = Object.freeze({ nested: { x: 1 } });
obj.nested.x = 99; // nested not frozen
console.log(obj.nested.x);
"#), vec!["99"]);
}

#[test]
fn deep_freeze_pattern() {
    assert_eq!(run_js(r#"
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
"#), vec!["5432"]);
}

#[test]
fn frozen_array_prevents_push() {
    assert_eq!(run_js(r#"
const arr = Object.freeze([1, 2, 3]);
let threw = false;
try { arr.push(4); } catch { threw = true; }
console.log(threw);
console.log(arr.length);
"#), vec!["true", "3"]);
}

#[test]
fn seal_allows_modify_prevents_add_delete() {
    assert_eq!(run_js(r#"
const obj = Object.seal({ x: 1, y: 2 });
obj.x = 99; // modify ok
obj.z = 3;  // add — silently fails
delete obj.y; // delete — silently fails
console.log(obj.x);
console.log(obj.z);
console.log(obj.y);
"#), vec!["99", "undefined", "2"]);
}

#[test]
fn is_frozen_empty_non_extensible_is_frozen() {
    assert_eq!(run_js(r#"
const obj = Object.preventExtensions({});
console.log(Object.isFrozen(obj)); // empty + non-extensible = frozen
"#), vec!["true"]);
}

#[test]
fn object_create_null_no_prototype() {
    assert_eq!(run_js(r#"
const safe = Object.create(null);
safe.key = "value";
// No toString, no hasOwnProperty — truly property-free
console.log(typeof safe.toString);
console.log(safe.key);
"#), vec!["undefined", "value"]);
}

#[test]
fn prevent_extensions_allows_modify() {
    assert_eq!(run_js(r#"
const obj = Object.preventExtensions({ x: 1 });
obj.x = 99;  // existing properties can be modified
obj.y = 2;   // new properties silently fail
console.log(obj.x);
console.log(obj.y);
"#), vec!["99", "undefined"]);
}
