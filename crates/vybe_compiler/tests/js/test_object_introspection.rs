use super::helpers::run_js;

// ===================================================================
// 1. Object.getOwnPropertyNames — includes non-enumerable properties
// ===================================================================

#[test]
fn get_own_property_names_includes_non_enumerable() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", { value: 99, enumerable: false });
const names = Object.getOwnPropertyNames(obj);
console.log(names.includes("a"));
console.log(names.includes("b"));
console.log(names.includes("hidden"));
"#), vec!["true", "true", "true"]);
}

// ===================================================================
// 2. Object.getOwnPropertyDescriptor — returns a descriptor object with value
// ===================================================================

#[test]
fn get_own_property_descriptor_all_flags() {
    assert_eq!(run_js(r#"
const obj = { x: 42 };
const d = Object.getOwnPropertyDescriptor(obj, "x");
console.log(d.value);
console.log(d.writable);
console.log(d.enumerable);
console.log(d.configurable);
"#), vec!["42", "true", "true", "true"]);
}

// ===================================================================
// 3. Object.getOwnPropertyDescriptors — all descriptors at once
// ===================================================================

#[test]
fn get_own_property_descriptors_returns_all() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
Object.defineProperty(obj, "c", { value: 3, enumerable: false, writable: true, configurable: true });
const descs = Object.getOwnPropertyDescriptors(obj);
console.log(descs.a.value);
console.log(descs.b.value);
console.log(descs.c.value);
console.log(descs.c.enumerable);
"#), vec!["1", "2", "3", "false"]);
}

// ===================================================================
// 4. Object.getOwnPropertySymbols — returns symbol-keyed properties
// ===================================================================

#[test]
fn get_own_property_symbols_returns_symbol_keys() {
    assert_eq!(run_js(r#"
const sym = Symbol("tag");
const obj = { normal: 1, [sym]: "symbolValue" };
const syms = Object.getOwnPropertySymbols(obj);
console.log(syms.length);
console.log(obj[syms[0]]);
"#), vec!["1", "symbolValue"]);
}

// ===================================================================
// 5. defineProperty enumerable:false — doesn't appear in for..in
// ===================================================================

#[test]
fn define_property_non_enumerable_absent_from_for_in() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
Object.defineProperty(obj, "secret", { value: 42, enumerable: false });
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("a"));
console.log(keys.includes("secret"));
console.log(obj.secret);
"#), vec!["true", "false", "42"]);
}

// ===================================================================
// 6. defineProperty writable:false — assignment silently ignored
// ===================================================================

#[test]
fn define_property_writable_false_ignores_assignment() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "CONST", { value: 100, writable: false, enumerable: true, configurable: false });
obj.CONST = 999;
console.log(obj.CONST);
"#), vec!["100"]);
}

// ===================================================================
// 7. defineProperty configurable:false — cannot redefine
// ===================================================================

#[test]
fn define_property_configurable_false_prevents_redefine() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "locked", { value: 1, configurable: false, writable: false, enumerable: true });
let threw = false;
try {
    Object.defineProperty(obj, "locked", { value: 2 });
} catch (e) {
    threw = true;
}
console.log(threw);
console.log(obj.locked);
"#), vec!["true", "1"]);
}

// ===================================================================
// 8. defineProperty getter/setter pair
// ===================================================================

#[test]
fn define_property_getter_setter_pair() {
    assert_eq!(run_js(r#"
const obj = { _count: 0 };
Object.defineProperty(obj, "count", {
    get() { return this._count; },
    set(v) { this._count = v < 0 ? 0 : v; },
    enumerable: true,
    configurable: true
});
obj.count = 5;
console.log(obj.count);
obj.count = -3;
console.log(obj.count);
"#), vec!["5", "0"]);
}

// ===================================================================
// 9. Object.defineProperties — multiple properties at once
// ===================================================================

#[test]
fn define_properties_multiple_at_once() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperties(obj, {
    x: { value: 10, enumerable: true, writable: true, configurable: true },
    y: { value: 20, enumerable: true, writable: true, configurable: true },
    z: { value: 30, enumerable: false, writable: false, configurable: false }
});
console.log(obj.x);
console.log(obj.y);
console.log(obj.z);
console.log(Object.keys(obj).join(","));
"#), vec!["10", "20", "30", "x,y"]);
}

// ===================================================================
// 10. Object.isFrozen — true after freeze
// ===================================================================

#[test]
fn is_frozen_true_after_freeze() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
Object.freeze(obj);
console.log(Object.isFrozen(obj));
"#), vec!["true"]);
}

// ===================================================================
// 11. Object.isFrozen — false on normal mutable object
// ===================================================================

#[test]
fn is_frozen_false_on_normal_object() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
console.log(Object.isFrozen(obj));
"#), vec!["false"]);
}

// ===================================================================
// 12. Object.isSealed — true after seal
// ===================================================================

#[test]
fn is_sealed_true_after_seal() {
    assert_eq!(run_js(r#"
const obj = { x: 10 };
Object.seal(obj);
console.log(Object.isSealed(obj));
"#), vec!["true"]);
}

// ===================================================================
// 13. Object.isSealed — false on normal object
// ===================================================================

#[test]
fn is_sealed_false_on_normal_object() {
    assert_eq!(run_js(r#"
const obj = { x: 10 };
console.log(Object.isSealed(obj));
"#), vec!["false"]);
}

// ===================================================================
// 14. Object.isExtensible — false after preventExtensions
// ===================================================================

#[test]
fn is_extensible_false_after_prevent_extensions() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
console.log(Object.isExtensible(obj));
Object.preventExtensions(obj);
console.log(Object.isExtensible(obj));
"#), vec!["true", "false"]);
}

// ===================================================================
// 15. Object.preventExtensions — blocks adding new properties
// ===================================================================

#[test]
fn prevent_extensions_blocks_new_properties() {
    assert_eq!(run_js(r#"
const obj = { existing: "yes" };
Object.preventExtensions(obj);
obj.newProp = "no";
console.log(obj.existing);
console.log(obj.newProp);
"#), vec!["yes", "undefined"]);
}

// ===================================================================
// 16. Object.seal — can modify existing, can't add or delete
// ===================================================================

#[test]
fn seal_allows_modify_blocks_add_and_delete() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
Object.seal(obj);
obj.a = 99;
obj.c = 3;
delete obj.b;
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#), vec!["99", "2", "undefined"]);
}

// ===================================================================
// 17. Object.freeze — can't add, delete, or modify
// ===================================================================

#[test]
fn freeze_blocks_add_delete_and_modify() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2 };
Object.freeze(obj);
obj.a = 99;
obj.c = 3;
delete obj.b;
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#), vec!["1", "2", "undefined"]);
}

// ===================================================================
// 18. Object.getPrototypeOf — returns the prototype
// ===================================================================

#[test]
fn get_prototype_of_returns_prototype() {
    assert_eq!(run_js(r#"
const proto = { kind: "proto" };
const obj = Object.create(proto);
console.log(Object.getPrototypeOf(obj) === proto);
console.log(obj.kind);
"#), vec!["true", "proto"]);
}

// ===================================================================
// 19. Object.setPrototypeOf — changes prototype chain
// ===================================================================

#[test]
fn set_prototype_of_changes_prototype() {
    assert_eq!(run_js(r#"
const newProto = { greet() { return "hello from newProto"; } };
const obj = { x: 1 };
Object.setPrototypeOf(obj, newProto);
console.log(Object.getPrototypeOf(obj) === newProto);
console.log(obj.greet());
"#), vec!["true", "hello from newProto"]);
}

// ===================================================================
// 20. Object.create with property descriptors
// ===================================================================

#[test]
fn object_create_with_descriptors() {
    assert_eq!(run_js(r#"
const obj = Object.create(Object.prototype, {
    name: { value: "Alice", enumerable: true, writable: true, configurable: true },
    age:  { value: 30,      enumerable: true, writable: false, configurable: false }
});
console.log(obj.name);
console.log(obj.age);
obj.age = 99;
console.log(obj.age);
"#), vec!["Alice", "30", "30"]);
}

// ===================================================================
// 21. Object.create(null) — no prototype
// ===================================================================

#[test]
fn object_create_null_no_prototype() {
    assert_eq!(run_js(r#"
const obj = Object.create(null);
obj.foo = "bar";
console.log(obj.foo);
console.log(Object.getPrototypeOf(obj));
"#), vec!["bar", "null"]);
}

// ===================================================================
// 22. Object.assign — copies only own enumerable properties
// ===================================================================

#[test]
fn object_assign_copies_own_enumerable_only() {
    assert_eq!(run_js(r#"
const src = { a: 1, b: 2 };
Object.defineProperty(src, "hidden", { value: 99, enumerable: false });
const dest = {};
Object.assign(dest, src);
console.log(dest.a);
console.log(dest.b);
console.log(dest.hidden);
"#), vec!["1", "2", "undefined"]);
}

// ===================================================================
// 23. Object.entries — returns [key, value] pairs
// ===================================================================

#[test]
fn object_entries_returns_key_value_pairs() {
    assert_eq!(run_js(r#"
const obj = { x: 10, y: 20, z: 30 };
const entries = Object.entries(obj);
entries.forEach(([k, v]) => console.log(k + "=" + v));
"#), vec!["x=10", "y=20", "z=30"]);
}

// ===================================================================
// 24. Object.fromEntries — from array of pairs
// ===================================================================

#[test]
fn object_from_entries_from_array() {
    assert_eq!(run_js(r#"
const pairs = [["name", "Bob"], ["age", 25], ["city", "Paris"]];
const obj = Object.fromEntries(pairs);
console.log(obj.name);
console.log(obj.age);
console.log(obj.city);
"#), vec!["Bob", "25", "Paris"]);
}

// ===================================================================
// 25. Object.fromEntries — from Map
// ===================================================================

#[test]
fn object_from_entries_from_map() {
    assert_eq!(run_js(r#"
const map = new Map([["one", 1], ["two", 2], ["three", 3]]);
const obj = Object.fromEntries(map);
console.log(obj.one);
console.log(obj.two);
console.log(obj.three);
"#), vec!["1", "2", "3"]);
}

// ===================================================================
// 26. hasOwnProperty — own vs inherited property distinction
// ===================================================================

#[test]
fn has_own_property_vs_inherited() {
    assert_eq!(run_js(r#"
const proto = { inherited: true };
const obj = Object.create(proto);
obj.own = true;
console.log(obj.hasOwnProperty("own"));
console.log(obj.hasOwnProperty("inherited"));
console.log(obj.inherited);
"#), vec!["true", "false", "true"]);
}

// ===================================================================
// 27. propertyIsEnumerable — own enumerable vs inherited
// ===================================================================

#[test]
fn property_is_enumerable_own_vs_inherited() {
    assert_eq!(run_js(r#"
const proto = { fromProto: 1 };
const obj = Object.create(proto);
obj.ownEnum = 2;
Object.defineProperty(obj, "ownHidden", { value: 3, enumerable: false });
console.log(obj.propertyIsEnumerable("ownEnum"));
console.log(obj.propertyIsEnumerable("ownHidden"));
console.log(obj.propertyIsEnumerable("fromProto"));
"#), vec!["true", "false", "false"]);
}

// ===================================================================
// 28. Object.keys — only own enumerable properties
// ===================================================================

#[test]
fn object_keys_own_enumerable_only() {
    assert_eq!(run_js(r#"
const proto = { inherited: 0 };
const obj = Object.create(proto);
obj.a = 1;
obj.b = 2;
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Object.keys(obj);
console.log(keys.join(","));
console.log(keys.includes("inherited"));
console.log(keys.includes("hidden"));
"#), vec!["a,b", "false", "false"]);
}

// ===================================================================
// 29. Object.values — only own enumerable values
// ===================================================================

#[test]
fn object_values_own_enumerable_only() {
    assert_eq!(run_js(r#"
const obj = { p: 10, q: 20 };
Object.defineProperty(obj, "secret", { value: 99, enumerable: false });
const vals = Object.values(obj);
console.log(vals.join(","));
console.log(vals.includes(99));
"#), vec!["10,20", "false"]);
}

// ===================================================================
// 30. toString override via prototype
// ===================================================================

#[test]
fn tostring_override_via_prototype() {
    assert_eq!(run_js(r#"
function Point(x, y) {
    this.x = x;
    this.y = y;
}
Point.prototype.toString = function() {
    return "(" + this.x + "," + this.y + ")";
};
const p = new Point(3, 4);
console.log(p.toString());
console.log(String(p));
"#), vec!["(3,4)", "(3,4)"]);
}
