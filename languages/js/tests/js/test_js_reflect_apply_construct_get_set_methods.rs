use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Reflect Object Methods (apply, construct, get, set, deleteProperty, has, isExtensible)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_reflect_apply_invocation() {
    let src = r#"
function greet(prefix, suffix) {
    return `${prefix} ${this.name} ${suffix}`;
}
const result = Reflect.apply(greet, { name: "World" }, ["Hello", "!"]);
console.log(result);
"#;
    assert_eq!(run_js(src), vec!["Hello World !"]);
}

#[test]
fn test_js_reflect_construct_invocation() {
    let src = r#"
function Person(name, age) {
    this.name = name;
    this.age = age;
}
const p = Reflect.construct(Person, ["Alice", 30]);
console.log(p.name + "|" + p.age + "|" + (p instanceof Person));
"#;
    assert_eq!(run_js(src), vec!["Alice|30|true"]);
}

#[test]
fn test_js_reflect_construct_with_new_target_override() {
    let src = r#"
function Base() { this.base = true; }
function Sub() {}
Sub.prototype = Object.create(Base.prototype);
Sub.prototype.subProp = "sub";

const obj = Reflect.construct(Base, [], Sub);
console.log(obj.subProp + "|" + (obj instanceof Sub));
"#;
    assert_eq!(run_js(src), vec!["sub|true"]);
}

#[test]
fn test_js_reflect_get_property_value() {
    let src = r#"
const obj = { x: 10, y: 20 };
console.log(Reflect.get(obj, "x") + "|" + Reflect.get(obj, "y"));
"#;
    assert_eq!(run_js(src), vec!["10|20"]);
}

#[test]
fn test_js_reflect_get_with_custom_receiver_this() {
    let src = r#"
const target = {
    _val: 100,
    get val() { return this._val; }
};
const receiver = { _val: 999 };
console.log(Reflect.get(target, "val", receiver));
"#;
    assert_eq!(run_js(src), vec!["999"]);
}

#[test]
fn test_js_reflect_set_property_returns_boolean() {
    let src = r#"
const obj = {};
const success = Reflect.set(obj, "prop", 42);
console.log(success + "|" + obj.prop);
"#;
    assert_eq!(run_js(src), vec!["true|42"]);
}

#[test]
fn test_js_reflect_set_non_writable_returns_false() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 10, writable: false });
const success = Reflect.set(obj, "fixed", 99);
console.log(success + "|" + obj.fixed);
"#;
    assert_eq!(run_js(src), vec!["false|10"]);
}

#[test]
fn test_js_reflect_has_in_operator_equivalent() {
    let src = r#"
const proto = { inherited: true };
const obj = Object.create(proto);
obj.own = true;
console.log(Reflect.has(obj, "own") + "|" + Reflect.has(obj, "inherited") + "|" + Reflect.has(obj, "missing"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_reflect_delete_property_returns_boolean() {
    let src = r#"
const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, configurable: false });
console.log(Reflect.deleteProperty(obj, "a"));
console.log(Reflect.deleteProperty(obj, "b"));
console.log("a" in obj);
"#;
    assert_eq!(run_js(src), vec!["true", "false", "false"]);
}

#[test]
fn test_js_reflect_get_prototype_of() {
    let src = r#"
const proto = { val: 1 };
const obj = Object.create(proto);
console.log(Reflect.getPrototypeOf(obj) === proto);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_reflect_set_prototype_of_returns_boolean() {
    let src = r#"
const obj = {};
const proto = { newProto: true };
const success = Reflect.setPrototypeOf(obj, proto);
console.log(success + "|" + obj.newProto);
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_reflect_set_prototype_of_cycle_returns_false() {
    let src = r#"
const a = {};
const b = Object.create(a);
console.log(Reflect.setPrototypeOf(a, b)); // Prototype cycle rejected!
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_reflect_prevent_extensions_returns_boolean() {
    let src = r#"
const obj = { a: 1 };
console.log(Reflect.preventExtensions(obj));
console.log(Reflect.isExtensible(obj));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_reflect_define_property_returns_boolean() {
    let src = r#"
const obj = {};
const success = Reflect.defineProperty(obj, "a", { value: 10, writable: true });
console.log(success + "|" + obj.a);
"#;
    assert_eq!(run_js(src), vec!["true|10"]);
}

#[test]
fn test_js_reflect_define_property_failure_returns_false_instead_of_throwing() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "locked", { value: 1, configurable: false });
const success = Reflect.defineProperty(obj, "locked", { configurable: true });
console.log(success); // Returns false cleanly without throwing exception!
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_reflect_own_keys_strings_and_symbols() {
    let src = r#"
const sym = Symbol("s");
const obj = { b: 2, a: 1, [sym]: 3 };
const keys = Reflect.ownKeys(obj);
console.log(keys.length + "|" + (keys[2] === sym));
"#;
    assert_eq!(run_js(src), vec!["3|true"]);
}

#[test]
fn test_js_reflect_get_own_property_descriptor() {
    let src = r#"
const obj = { x: 100 };
const desc = Reflect.getOwnPropertyDescriptor(obj, "x");
console.log(desc.value + "|" + desc.writable);
"#;
    assert_eq!(run_js(src), vec!["100|true"]);
}

#[test]
fn test_js_reflect_apply_on_builtins() {
    let src = r#"
const str = "hello";
const result = Reflect.apply(String.prototype.toUpperCase, str, []);
console.log(result);
"#;
    assert_eq!(run_js(src), vec!["HELLO"]);
}

#[test]
fn test_js_reflect_set_with_setter_receiver() {
    let src = r#"
const target = {
    set val(v) { this._val = v * 2; }
};
const receiver = {};
Reflect.set(target, "val", 10, receiver);
console.log(receiver._val);
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_reflect_is_extensible_primitives_throw_typeerror() {
    let src = r#"
try {
    Reflect.isExtensible(42);
} catch (e) {
    console.log("Reflect Primitive TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect Primitive TypeError"]);
}
