use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Prototype Chain, Property Shadowing & Lookup Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_prototype_chain_shadowing_instance_property() {
    let src = r#"
const proto = { value: "PrototypeVal" };
const obj = Object.create(proto);
console.log(obj.value);
obj.value = "ShadowedVal";
console.log(obj.value + "|" + proto.value);
"#;
    assert_eq!(
        run_js(src),
        vec!["PrototypeVal", "ShadowedVal|PrototypeVal"]
    );
}

#[test]
fn test_js_prototype_chain_non_writable_prototype_property_shadowing_blocked() {
    let src = r#"
const proto = {};
Object.defineProperty(proto, "fixed", { value: 10, writable: false });
const obj = Object.create(proto);

try {
    "use strict";
    obj.fixed = 20; // Cannot shadow non-writable property on prototype via normal assignment!
} catch (e) {
    console.log("Shadowing Non-Writable Prototype Property TypeError");
}
console.log(obj.fixed);
"#;
    assert_eq!(
        run_js(src),
        vec!["Shadowing Non-Writable Prototype Property TypeError", "10"]
    );
}

#[test]
fn test_js_prototype_chain_define_property_bypasses_non_writable_prototype() {
    let src = r#"
const proto = {};
Object.defineProperty(proto, "fixed", { value: 10, writable: false });
const obj = Object.create(proto);

// Object.defineProperty directly defines own property, bypassing prototype check!
Object.defineProperty(obj, "fixed", { value: 20, writable: true });
console.log(obj.fixed + "|" + proto.fixed);
"#;
    assert_eq!(run_js(src), vec!["20|10"]);
}

#[test]
fn test_js_prototype_chain_setter_on_prototype_intercepts_assignment() {
    let src = r#"
const proto = {
    set count(v) { this._count = v * 10; },
    get count() { return this._count; }
};
const obj = Object.create(proto);
obj.count = 5; // Triggers prototype setter with 'this' pointing to obj!
console.log(obj._count + "|" + (Object.hasOwn(obj, "count")));
"#;
    assert_eq!(run_js(src), vec!["50|false"]);
}

#[test]
fn test_js_prototype_chain_multi_level_lookup() {
    let src = r#"
const g = { level: "GrandParent" };
const p = Object.create(g);
const c = Object.create(p);

console.log(c.level);
p.level = "Parent";
console.log(c.level);
c.level = "Child";
console.log(c.level + "|" + p.level + "|" + g.level);
"#;
    assert_eq!(
        run_js(src),
        vec!["GrandParent", "Parent", "Child|Parent|GrandParent"]
    );
}

#[test]
fn test_js_prototype_chain_delete_reveals_prototype_property() {
    let src = r#"
const proto = { item: "ProtoItem" };
const obj = Object.create(proto);
obj.item = "OwnItem";
console.log(obj.item);
delete obj.item;
console.log(obj.item);
"#;
    assert_eq!(run_js(src), vec!["OwnItem", "ProtoItem"]);
}

#[test]
fn test_js_prototype_chain_in_operator_traverses_chain() {
    let src = r#"
const proto = { key: 1 };
const obj = Object.create(proto);
console.log(("key" in obj) + "|" + Object.hasOwn(obj, "key"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_prototype_chain_for_in_loop_traverses_enumerable_properties() {
    let src = r#"
const proto = { protoKey: 100 };
const obj = Object.create(proto);
obj.ownKey = 200;

const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["ownKey,protoKey"]);
}

#[test]
fn test_js_prototype_chain_for_in_loop_shadowing_hides_prototype_key() {
    let src = r#"
const proto = { key: "proto" };
const obj = Object.create(proto);
obj.key = "own";

const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(",") + "|Count=" + keys.length);
"#;
    assert_eq!(run_js(src), vec!["key|Count=1"]);
}

#[test]
fn test_js_prototype_chain_null_prototype_termination() {
    let src = r#"
const obj = Object.create(null);
console.log(Object.getPrototypeOf(obj) === null);
console.log(obj.toString === undefined);
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_prototype_chain_method_this_binding() {
    let src = r#"
const proto = {
    multiplier: 2,
    calc(x) { return x * this.multiplier; }
};
const obj = Object.create(proto);
obj.multiplier = 10;
console.log(obj.calc(5));
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_prototype_chain_dynamic_prototype_mutation() {
    let src = r#"
const proto1 = { v: "P1" };
const proto2 = { v: "P2" };
const obj = Object.create(proto1);
console.log(obj.v);
Object.setPrototypeOf(obj, proto2);
console.log(obj.v);
"#;
    assert_eq!(run_js(src), vec!["P1", "P2"]);
}

#[test]
fn test_js_prototype_chain_symbol_property_lookup() {
    let src = r#"
const sym = Symbol("sym");
const proto = { [sym]: "ProtoSymbolVal" };
const obj = Object.create(proto);
console.log(obj[sym]);
"#;
    assert_eq!(run_js(src), vec!["ProtoSymbolVal"]);
}

#[test]
fn test_js_prototype_chain_getter_without_setter_shadowing_fails_in_strict() {
    let src = r#"
const proto = {
    get readOnlyProp() { return "ReadOnly"; }
};
const obj = Object.create(proto);
try {
    "use strict";
    obj.readOnlyProp = "NewVal";
} catch (e) {
    console.log("Assign ReadOnly Prototype Getter TypeError");
}
console.log(obj.readOnlyProp);
"#;
    assert_eq!(
        run_js(src),
        vec!["Assign ReadOnly Prototype Getter TypeError", "ReadOnly"]
    );
}

#[test]
fn test_js_prototype_chain_object_keys_does_not_traverse_chain() {
    let src = r#"
const proto = { protoKey: 1 };
const obj = Object.create(proto);
obj.ownKey = 2;
console.log(Object.keys(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["ownKey"]);
}

#[test]
fn test_js_prototype_chain_function_prototype_constructor_reference() {
    let src = r#"
function Person(name) { this.name = name; }
const p = new Person("Alice");
console.log(p.constructor === Person);
console.log(Person.prototype.constructor === Person);
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_prototype_chain_constructor_shadowing() {
    let src = r#"
function Base() {}
const obj = new Base();
obj.constructor = "CustomConstructorString";
console.log(obj.constructor);
"#;
    assert_eq!(run_js(src), vec!["CustomConstructorString"]);
}

#[test]
fn test_js_prototype_chain_non_enumerable_prototype_properties() {
    let src = r#"
const proto = {};
Object.defineProperty(proto, "hidden", { value: 100, enumerable: false });
const obj = Object.create(proto);
console.log(obj.hidden + "|" + ("hidden" in obj));
"#;
    assert_eq!(run_js(src), vec!["100|true"]);
}

#[test]
fn test_js_prototype_chain_array_prototype_method_borrowing() {
    let src = r#"
const arrayLike = { 0: "a", 1: "b", length: 2 };
const joined = Array.prototype.join.call(arrayLike, "-");
console.log(joined);
"#;
    assert_eq!(run_js(src), vec!["a-b"]);
}

#[test]
fn test_js_prototype_chain_cycle_detection_throws_typeerror() {
    let src = r#"
const a = {};
const b = Object.create(a);
try {
    Object.setPrototypeOf(a, b);
} catch (e) {
    console.log("Prototype Cycle TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Prototype Cycle TypeError"]);
}
