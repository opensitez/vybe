/// Reflect API — Reflect.apply, Reflect.construct, Reflect.defineProperty,
/// Reflect.deleteProperty, Reflect.get/set, Reflect.has, Reflect.ownKeys,
/// Reflect.getPrototypeOf/setPrototypeOf, Reflect.isExtensible/preventExtensions.
use super::helpers::run_js;

#[test]
fn reflect_apply_calls_function() {
    assert_eq!(
        run_js(
            r#"
function sum(a, b) { return a + b; }
console.log(Reflect.apply(sum, null, [3, 4]));
"#
        ),
        vec!["7"]
    );
}

#[test]
fn reflect_apply_with_this() {
    assert_eq!(
        run_js(
            r#"
function greet() { return "Hello " + this.name; }
const obj = { name: "World" };
console.log(Reflect.apply(greet, obj, []));
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn reflect_construct_creates_instance() {
    assert_eq!(
        run_js(
            r#"
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
const p = Reflect.construct(Point, [3, 4]);
console.log(p.x);
console.log(p.y);
console.log(p instanceof Point);
"#
        ),
        vec!["3", "4", "true"]
    );
}

#[test]
fn reflect_get_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 42 };
console.log(Reflect.get(obj, "x"));
console.log(Reflect.get(obj, "missing"));
"#
        ),
        vec!["42", "undefined"]
    );
}

#[test]
fn reflect_set_property() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
const result = Reflect.set(obj, "x", 99);
console.log(result); // true on success
console.log(obj.x);
"#
        ),
        vec!["true", "99"]
    );
}

#[test]
fn reflect_has_checks_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
const proto = { inherited: true };
const obj = Object.create(proto);
obj.own = true;
console.log(Reflect.has(obj, "own"));
console.log(Reflect.has(obj, "inherited"));
console.log(Reflect.has(obj, "missing"));
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn reflect_delete_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
console.log(Reflect.deleteProperty(obj, "a"));
console.log("a" in obj);
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn reflect_delete_non_configurable_fails() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "x", { value: 1, configurable: false });
console.log(Reflect.deleteProperty(obj, "x")); // false
console.log("x" in obj);
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn reflect_own_keys_all_types() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("s");
const obj = { a: 1, [sym]: 2 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Reflect.ownKeys(obj);
console.log(keys.includes("a"));
console.log(keys.includes("hidden"));
console.log(keys.some(k => typeof k === "symbol"));
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn reflect_define_property() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Reflect.defineProperty(obj, "x", { value: 42, writable: false, enumerable: true, configurable: false });
console.log(obj.x);
obj.x = 99; // silently fails — not writable
console.log(obj.x);
"#
        ),
        vec!["42", "42"]
    );
}

#[test]
fn reflect_get_prototype_of() {
    assert_eq!(
        run_js(
            r#"
class Foo {}
const f = new Foo();
console.log(Reflect.getPrototypeOf(f) === Foo.prototype);
console.log(Reflect.getPrototypeOf(Foo.prototype) === Object.prototype);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn reflect_set_prototype_of() {
    assert_eq!(
        run_js(
            r#"
const a = { hello() { return "a"; } };
const b = { hello() { return "b"; } };
const obj = Object.create(a);
console.log(obj.hello());
Reflect.setPrototypeOf(obj, b);
console.log(obj.hello());
"#
        ),
        vec!["a", "b"]
    );
}

#[test]
fn reflect_is_extensible() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
console.log(Reflect.isExtensible(obj));
Reflect.preventExtensions(obj);
console.log(Reflect.isExtensible(obj));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn reflect_prevent_extensions_blocks_new_props() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
Reflect.preventExtensions(obj);
obj.y = 2; // silently fails
console.log("y" in obj);
obj.x = 99; // existing props still modifiable
console.log(obj.x);
"#
        ),
        vec!["false", "99"]
    );
}
