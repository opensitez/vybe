use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Reflect API Methods (`apply`, `construct`, `get`, `set`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_reflect_apply_basic_invocation() {
    let src = r#"
function add(a, b) { return a + b + this.bonus; }
const ctx = { bonus: 10 };
console.log(Reflect.apply(add, ctx, [5, 15]));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_reflect_construct_basic_instantiation() {
    let src = r#"
class Person {
    constructor(name, age) {
        this.name = name;
        this.age = age;
    }
}
const p = Reflect.construct(Person, ["Alice", 30]);
console.log(`${p.name}:${p.age}|isPerson=${p instanceof Person}`);
"#;
    assert_eq!(run_js(src), vec!["Alice:30|isPerson=true"]);
}

#[test]
fn test_js_reflect_construct_new_target_override() {
    let src = r#"
class Base {
    constructor() {
        this.targetName = new.target.name;
    }
}
class CustomTarget {}

const obj = Reflect.construct(Base, [], CustomTarget);
console.log(obj.targetName + "|isCustom=" + (obj instanceof CustomTarget));
"#;
    assert_eq!(run_js(src), vec!["CustomTarget|isCustom=true"]);
}

#[test]
fn test_js_reflect_get_property_lookup() {
    let src = r#"
const obj = { x: 10, y: 20 };
console.log(Reflect.get(obj, "x") + "|" + Reflect.get(obj, "missing"));
"#;
    assert_eq!(run_js(src), vec!["10|undefined"]);
}

#[test]
fn test_js_reflect_get_with_receiver_override() {
    let src = r#"
const proto = {
    get val() { return this._val; }
};
const receiver = { _val: "ReceiverValue" };
console.log(Reflect.get(proto, "val", receiver));
"#;
    assert_eq!(run_js(src), vec!["ReceiverValue"]);
}

#[test]
fn test_js_reflect_set_property_assignment() {
    let src = r#"
const obj = { x: 1 };
const success = Reflect.set(obj, "x", 99);
console.log(success + "|" + obj.x);
"#;
    assert_eq!(run_js(src), vec!["true|99"]);
}

#[test]
fn test_js_reflect_set_with_receiver_invokes_setter_on_receiver() {
    let src = r#"
const proto = {
    set score(v) { this._score = v * 2; }
};
const receiver = {};
const success = Reflect.set(proto, "score", 50, receiver);
console.log(success + "|" + receiver._score);
"#;
    assert_eq!(run_js(src), vec!["true|100"]);
}

#[test]
fn test_js_reflect_set_non_writable_property_returns_false() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 10, writable: false });
const success = Reflect.set(obj, "fixed", 20);
console.log(success + "|" + obj.fixed);
"#;
    assert_eq!(run_js(src), vec!["false|10"]);
}

#[test]
fn test_js_reflect_apply_non_callable_throws_typeerror() {
    let src = r#"
try {
    Reflect.apply("not_a_fn", null, []);
} catch (e) {
    console.log("Reflect.apply Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect.apply Non-Callable TypeError"]);
}

#[test]
fn test_js_reflect_construct_non_constructor_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.construct(() => {}, []);
} catch (e) {
    console.log("Reflect.construct Non-Constructor TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Reflect.construct Non-Constructor TypeError"]
    );
}

#[test]
fn test_js_reflect_get_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.get(123, "prop");
} catch (e) {
    console.log("Reflect.get Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect.get Non-Object TypeError"]);
}

#[test]
fn test_js_reflect_set_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.set("str", "prop", "val");
} catch (e) {
    console.log("Reflect.set Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect.set Non-Object TypeError"]);
}

#[test]
fn test_js_reflect_get_symbol_property() {
    let src = r#"
const sym = Symbol("key");
const obj = { [sym]: "SymbolData" };
console.log(Reflect.get(obj, sym));
"#;
    assert_eq!(run_js(src), vec!["SymbolData"]);
}

#[test]
fn test_js_reflect_set_symbol_property() {
    let src = r#"
const sym = Symbol("key");
const obj = {};
Reflect.set(obj, sym, "NewSymbolData");
console.log(obj[sym]);
"#;
    assert_eq!(run_js(src), vec!["NewSymbolData"]);
}

#[test]
fn test_js_reflect_get_inherited_prototype_property() {
    let src = r#"
const proto = { parentProp: "ParentValue" };
const obj = Object.create(proto);
console.log(Reflect.get(obj, "parentProp"));
"#;
    assert_eq!(run_js(src), vec!["ParentValue"]);
}

#[test]
fn test_js_reflect_apply_array_like_arguments() {
    let src = r#"
function fn(a, b) { return a + b; }
const args = { 0: 10, 1: 20, length: 2 };
console.log(Reflect.apply(fn, null, args));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_reflect_construct_array_like_arguments() {
    let src = r#"
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
const args = { 0: 5, 1: 15, length: 2 };
const pt = Reflect.construct(Point, args);
console.log(`${pt.x},${pt.y}`);
"#;
    assert_eq!(run_js(src), vec!["5,15"]);
}

#[test]
fn test_js_reflect_set_non_extensible_object_returns_false() {
    let src = r#"
const obj = Object.preventExtensions({ a: 1 });
const success = Reflect.set(obj, "b", 2);
console.log(success + "|hasB=" + ("b" in obj));
"#;
    assert_eq!(run_js(src), vec!["false|hasB=false"]);
}

#[test]
fn test_js_reflect_get_array_index() {
    let src = r#"
const arr = [10, 20, 30];
console.log(Reflect.get(arr, 1) + "|" + Reflect.get(arr, "length"));
"#;
    assert_eq!(run_js(src), vec!["20|3"]);
}

#[test]
fn test_js_reflect_set_array_length_truncation() {
    let src = r#"
const arr = [1, 2, 3, 4];
Reflect.set(arr, "length", 2);
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}
