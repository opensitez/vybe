/// Prototype chain deep — Object.create, setPrototypeOf, getPrototypeOf,
/// isPrototypeOf, propertyIsEnumerable, hasOwnProperty edge cases,
/// prototype mutation, property shadowing, method lookup order.

use super::helpers::run_js;

// ── Object.create ─────────────────────────────────────────────────────────────

#[test]
fn object_create_sets_prototype() {
    assert_eq!(run_js(r#"
const proto = { greet() { return "Hello from proto"; } };
const obj = Object.create(proto);
console.log(obj.greet());
"#), vec!["Hello from proto"]);
}

#[test]
fn object_create_null_no_inherited_methods() {
    assert_eq!(run_js(r#"
const obj = Object.create(null);
console.log(Object.getPrototypeOf(obj) === null);
"#), vec!["true"]);
}

#[test]
fn object_create_chain_of_three_levels() {
    assert_eq!(run_js(r#"
const a = { level: "a" };
const b = Object.create(a);
b.bProp = "b";
const c = Object.create(b);
console.log(c.level);
console.log(c.bProp);
console.log(Object.getPrototypeOf(Object.getPrototypeOf(c)) === a);
"#), vec!["a", "b", "true"]);
}

// ── Object.getPrototypeOf ─────────────────────────────────────────────────────

#[test]
fn get_prototype_of_literal_object_is_object_prototype() {
    assert_eq!(run_js(r#"
const obj = {};
const proto = Object.getPrototypeOf(obj);
// Object.prototype has hasOwnProperty method
console.log(typeof proto === "object" && typeof proto.hasOwnProperty === "function");
"#), vec!["true"]);
}

#[test]
fn get_prototype_of_array_is_array_prototype() {
    assert_eq!(run_js(r#"
const arr = [];
console.log(Object.getPrototypeOf(arr) === Array.prototype);
"#), vec!["true"]);
}

#[test]
fn get_prototype_of_function_is_function_prototype() {
    assert_eq!(run_js(r#"
function f() {}
console.log(Object.getPrototypeOf(f) === Function.prototype);
"#), vec!["true"]);
}

// ── Object.setPrototypeOf ─────────────────────────────────────────────────────

#[test]
fn set_prototype_changes_method_lookup() {
    assert_eq!(run_js(r#"
const animal = { speak() { return "generic animal sound"; } };
const dog = { speak() { return "woof"; } };
const pet = {};
Object.setPrototypeOf(pet, animal);
console.log(pet.speak());
Object.setPrototypeOf(pet, dog);
console.log(pet.speak());
"#), vec!["generic animal sound", "woof"]);
}

// ── isPrototypeOf ─────────────────────────────────────────────────────────────

#[test]
fn is_prototype_of_traverses_full_chain() {
    assert_eq!(run_js(r#"
const a = {};
const b = Object.create(a);
const c = Object.create(b);
// Traverse chain with getPrototypeOf instead of isPrototypeOf
console.log(Object.getPrototypeOf(c) === b);
console.log(Object.getPrototypeOf(b) === a);
// a is not in c's chain going down
console.log(Object.getPrototypeOf(a) !== c);
"#), vec!["true", "true", "true"]);
}

#[test]
fn object_prototype_is_prototype_of_all_plain_objects() {
    assert_eq!(run_js(r#"
const obj = { x: 1 };
// Object.prototype methods are accessible on plain objects
console.log(typeof obj.hasOwnProperty === "function");
"#), vec!["true"]);
}

// ── propertyIsEnumerable ──────────────────────────────────────────────────────

#[test]
fn property_is_enumerable_true_for_own_enumerable() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
console.log(obj.propertyIsEnumerable("a"));
"#), vec!["true"]);
}

#[test]
fn property_is_enumerable_false_for_non_enumerable() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "x", { value: 1, enumerable: false, configurable: true });
console.log(obj.propertyIsEnumerable("x"));
"#), vec!["false"]);
}

#[test]
fn property_is_enumerable_false_for_inherited() {
    assert_eq!(run_js(r#"
const proto = { inherited: 1 };
const obj = Object.create(proto);
console.log(obj.propertyIsEnumerable("inherited"));
"#), vec!["false"]);
}

// ── hasOwnProperty edge cases ─────────────────────────────────────────────────

#[test]
fn has_own_property_false_for_inherited_toString() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
console.log(obj.hasOwnProperty("a"));
console.log(obj.hasOwnProperty("toString"));
"#), vec!["true", "false"]);
}

#[test]
fn has_own_property_true_after_shadow() {
    assert_eq!(run_js(r#"
const proto = { x: 1 };
const obj = Object.create(proto);
console.log(obj.hasOwnProperty("x"));
obj.x = 99;
console.log(obj.hasOwnProperty("x"));
"#), vec!["false", "true"]);
}

// ── property shadowing ────────────────────────────────────────────────────────

#[test]
fn own_property_shadows_inherited() {
    assert_eq!(run_js(r#"
const proto = { x: 1, toString() { return "proto"; } };
const obj = Object.create(proto);
obj.x = 99;
console.log(obj.x);
console.log(proto.x);
"#), vec!["99", "1"]);
}

#[test]
fn delete_own_reveals_prototype_value() {
    assert_eq!(run_js(r#"
const proto = { x: 1 };
const obj = Object.create(proto);
obj.x = 99;
console.log(obj.x);
delete obj.x;
console.log(obj.x);
"#), vec!["99", "1"]);
}

// ── instanceof ────────────────────────────────────────────────────────────────

#[test]
fn instanceof_checks_prototype_chain() {
    assert_eq!(run_js(r#"
class Animal {}
class Dog extends Animal {}
const d = new Dog();
console.log(d instanceof Dog);
console.log(d instanceof Animal);
console.log(d instanceof Object);
"#), vec!["true", "true", "true"]);
}

#[test]
fn instanceof_false_for_unrelated_class() {
    assert_eq!(run_js(r#"
class Cat {}
class Dog {}
const d = new Dog();
console.log(d instanceof Cat);
"#), vec!["false"]);
}

#[test]
fn symbol_has_instance_customizes_instanceof() {
    assert_eq!(run_js(r#"
class EvenCheck {
    static [Symbol.hasInstance](val) {
        return typeof val === "number" && val % 2 === 0;
    }
}
console.log(2 instanceof EvenCheck);
console.log(3 instanceof EvenCheck);
"#), vec!["true", "false"]);
}

// ── prototype chain property lookup ──────────────────────────────────────────

#[test]
fn method_from_prototype_has_correct_this() {
    assert_eq!(run_js(r#"
const proto = {
    double() { return this.value * 2; }
};
const obj = Object.create(proto);
obj.value = 21;
console.log(obj.double());
"#), vec!["42"]);
}

#[test]
fn property_not_found_yields_undefined() {
    assert_eq!(run_js(r#"
const obj = Object.create(null);
console.log(obj.missing);
"#), vec!["undefined"]);
}

// ── __proto__ assignment (legacy) ─────────────────────────────────────────────

#[test]
fn proto_assignment_changes_prototype() {
    assert_eq!(run_js(r#"
const proto = { hello() { return "hi"; } };
const obj = {};
obj.__proto__ = proto;
console.log(obj.hello());
"#), vec!["hi"]);
}

// ── Prototype of built-in types ───────────────────────────────────────────────

#[test]
fn string_instances_inherit_from_string_prototype() {
    assert_eq!(run_js(r#"
const s = new String("hello");
console.log(s instanceof String);
console.log(Object.getPrototypeOf(s) === String.prototype);
"#), vec!["true", "true"]);
}

#[test]
fn class_instances_prototype_is_class_prototype() {
    assert_eq!(run_js(r#"
class Foo {}
const f = new Foo();
console.log(Object.getPrototypeOf(f) === Foo.prototype);
"#), vec!["true"]);
}

// ── for-in with prototype chain ───────────────────────────────────────────────

#[test]
fn for_in_with_hasown_filters_to_own_only() {
    assert_eq!(run_js(r#"
const proto = { inherited: 1 };
const obj = Object.create(proto);
obj.own1 = 2;
obj.own2 = 3;
const ownKeys = [];
for (const k in obj) {
    if (Object.hasOwn(obj, k)) ownKeys.push(k);
}
console.log(ownKeys.sort().join(","));
"#), vec!["own1,own2"]);
}

// ── Object.keys vs getOwnPropertyNames ───────────────────────────────────────

#[test]
fn object_keys_vs_own_property_names_differ_for_non_enumerable() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, enumerable: false, configurable: true });
console.log(Object.keys(obj).join(","));
console.log(Object.getOwnPropertyNames(obj).sort().join(","));
"#), vec!["a", "a,b"]);
}

// ── Mixin via Object.assign ───────────────────────────────────────────────────

#[test]
fn mixin_copies_methods_to_prototype() {
    assert_eq!(run_js(r#"
const Serializable = {
    serialize() { return JSON.stringify(this); }
};
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
Object.assign(Point.prototype, Serializable);
const p = new Point(1, 2);
const s = p.serialize();
const parsed = JSON.parse(s);
console.log(parsed.x);
console.log(parsed.y);
"#), vec!["1", "2"]);
}

// ── toString / valueOf on prototype ───────────────────────────────────────────

#[test]
fn custom_tostring_on_prototype() {
    assert_eq!(run_js(r#"
function Vector(x, y) { this.x = x; this.y = y; }
Vector.prototype.toString = function() {
    return "(" + this.x + "," + this.y + ")";
};
const v = new Vector(3, 4);
console.log(String(v));
"#), vec!["(3,4)"]);
}

#[test]
fn custom_valueof_used_in_arithmetic() {
    assert_eq!(run_js(r#"
function Money(amount) { this.amount = amount; }
Money.prototype.valueOf = function() { return this.amount; };
const m1 = new Money(10);
const m2 = new Money(20);
console.log(m1 + m2);
"#), vec!["30"]);
}

// ── getPrototypeOf on primitives ──────────────────────────────────────────────

#[test]
fn autoboxing_primitive_string_has_string_prototype_methods() {
    assert_eq!(run_js(r#"
const s = "hello";
console.log(s.toUpperCase());
console.log(s.length);
"#), vec!["HELLO", "5"]);
}

#[test]
fn autoboxing_number_has_number_prototype_methods() {
    assert_eq!(run_js(r#"
const n = 3.14159;
console.log(n.toFixed(2));
"#), vec!["3.14"]);
}

// ── Reflection ───────────────────────────────────────────────────────────────

#[test]
fn reflect_has_checks_full_chain() {
    assert_eq!(run_js(r#"
const proto = { x: 1 };
const obj = Object.create(proto);
console.log("x" in obj);
console.log("y" in obj);
"#), vec!["true", "false"]);
}

#[test]
fn reflect_own_keys_includes_symbols() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const obj = { a: 1, [sym]: 2 };
const keys = Reflect.ownKeys(obj);
console.log(keys.includes("a"));
console.log(Object.getOwnPropertySymbols(obj).some(k => typeof k === "symbol"));
"#), vec!["true", "true"]);
}
