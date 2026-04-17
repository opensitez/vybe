/// JavaScript scope edge cases and prototype patterns:
/// temporal dead zone, block scoping, hoisting, prototype chain,
/// WeakMap, WeakSet, Symbol basics, Proxy/Reflect basics,
/// class fields, private methods, static blocks.

use super::helpers::run_js;

// ===================================================================
// SCOPE: BLOCK SCOPING WITH LET/CONST
// ===================================================================

#[test]
fn let_block_scope_in_if() {
    assert_eq!(run_js(r#"
let x = "outer";
if (true) {
    let x = "inner";
    console.log(x);
}
console.log(x);
"#), &["inner", "outer"]);
}

#[test]
fn const_block_scope() {
    assert_eq!(run_js(r#"
const x = 1;
{
    const x = 2;
    console.log(x);
}
console.log(x);
"#), &["2", "1"]);
}

#[test]
fn var_no_block_scope() {
    assert_eq!(run_js(r#"
var x = "outer";
if (true) {
    var x = "inner";
}
console.log(x);
"#), &["inner"]);
}

#[test]
fn let_in_for_loop_scope() {
    assert_eq!(run_js(r#"
let results = [];
for (let i = 0; i < 3; i++) {
    results.push(function() { return i; });
}
console.log(results[0]());
console.log(results[1]());
console.log(results[2]());
"#), &["0", "1", "2"]);
}

#[test]
fn var_in_for_loop_shared() {
    assert_eq!(run_js(r#"
var results = [];
for (var i = 0; i < 3; i++) {
    results.push(function() { return i; });
}
console.log(results[0]());
console.log(results[1]());
console.log(results[2]());
"#), &["3", "3", "3"]);
}

// ===================================================================
// HOISTING
// ===================================================================

#[test]
fn var_hoisting() {
    assert_eq!(run_js(r#"
console.log(typeof x);
var x = 5;
console.log(x);
"#), &["undefined", "5"]);
}

#[test]
fn function_hoisting() {
    assert_eq!(run_js(r#"
console.log(greet());
function greet() { return "hello"; }
"#), &["hello"]);
}

// ===================================================================
// SYMBOL BASICS
// ===================================================================

#[test]
fn symbol_basic() {
    assert_eq!(run_js(r#"
let s1 = Symbol("desc");
let s2 = Symbol("desc");
console.log(typeof s1);
console.log(s1 === s2);
console.log(s1.toString());
"#), &["symbol", "false", "Symbol(desc)"]);
}

#[test]
fn symbol_as_property_key() {
    assert_eq!(run_js(r#"
let id = Symbol("id");
let obj = { [id]: 42, name: "test" };
console.log(obj[id]);
console.log(Object.keys(obj).join(","));
"#), &["42", "name"]);
}

#[test]
fn symbol_for_shared() {
    assert_eq!(run_js(r#"
let s1 = Symbol.for("app.id");
let s2 = Symbol.for("app.id");
console.log(s1 === s2);
"#), &["true"]);
}

// ===================================================================
// WEAKMAP
// ===================================================================

#[test]
fn weakmap_basic() {
    assert_eq!(run_js(r#"
let wm = new WeakMap();
let key = {};
wm.set(key, "value");
console.log(wm.has(key));
console.log(wm.get(key));
wm.delete(key);
console.log(wm.has(key));
"#), &["true", "value", "false"]);
}

// ===================================================================
// WEAKSET
// ===================================================================

#[test]
fn weakset_basic() {
    assert_eq!(run_js(r#"
let ws = new WeakSet();
let obj = {};
ws.add(obj);
console.log(ws.has(obj));
ws.delete(obj);
console.log(ws.has(obj));
"#), &["true", "false"]);
}

// ===================================================================
// PROXY BASICS
// ===================================================================

#[test]
fn proxy_get_trap() {
    assert_eq!(run_js(r#"
let handler = {
    get(target, prop) {
        return prop in target ? target[prop] : "default";
    }
};
let obj = new Proxy({ name: "Alice" }, handler);
console.log(obj.name);
console.log(obj.missing);
"#), &["Alice", "default"]);
}

#[test]
fn proxy_set_trap() {
    assert_eq!(run_js(r#"
let handler = {
    set(target, prop, value) {
        if (typeof value !== "number") {
            throw new TypeError("Expected number");
        }
        target[prop] = value;
        return true;
    }
};
let obj = new Proxy({}, handler);
obj.x = 42;
console.log(obj.x);
try {
    obj.y = "string";
} catch (e) {
    console.log(e.message);
}
"#), &["42", "Expected number"]);
}

#[test]
fn proxy_has_trap() {
    assert_eq!(run_js(r#"
let handler = {
    has(target, prop) {
        if (prop.startsWith("_")) return false;
        return prop in target;
    }
};
let obj = new Proxy({ _secret: 1, visible: 2 }, handler);
console.log("visible" in obj);
console.log("_secret" in obj);
"#), &["true", "false"]);
}

// ===================================================================
// REFLECT BASICS
// ===================================================================

#[test]
fn reflect_get_set() {
    assert_eq!(run_js(r#"
let obj = { x: 1 };
console.log(Reflect.get(obj, "x"));
Reflect.set(obj, "y", 2);
console.log(obj.y);
"#), &["1", "2"]);
}

#[test]
fn reflect_has() {
    assert_eq!(run_js(r#"
let obj = { a: 1 };
console.log(Reflect.has(obj, "a"));
console.log(Reflect.has(obj, "b"));
"#), &["true", "false"]);
}

// ===================================================================
// CLASS FIELDS (PUBLIC / PRIVATE / STATIC)
// ===================================================================

#[test]
fn class_public_field_initializer() {
    assert_eq!(run_js(r#"
class Counter {
    count = 0;
    increment() { this.count++; }
}
let c = new Counter();
c.increment();
c.increment();
console.log(c.count);
"#), &["2"]);
}

#[test]
fn class_static_field() {
    assert_eq!(run_js(r#"
class Config {
    static version = "1.0";
    static appName = "MyApp";
}
console.log(Config.version);
console.log(Config.appName);
"#), &["1.0", "MyApp"]);
}

#[test]
fn class_private_field_encapsulation() {
    assert_eq!(run_js(r#"
class Secret {
    #value;
    constructor(v) { this.#value = v; }
    reveal() { return this.#value; }
}
let s = new Secret(42);
console.log(s.reveal());
console.log(s.value);
"#), &["42", "undefined"]);
}

#[test]
fn class_private_method() {
    assert_eq!(run_js(r#"
class Processor {
    #transform(x) { return x * 2; }
    process(x) { return this.#transform(x) + 1; }
}
let p = new Processor();
console.log(p.process(5));
"#), &["11"]);
}

// ===================================================================
// PROTOTYPE CHAIN PATTERNS
// ===================================================================

#[test]
fn prototype_method_lookup() {
    assert_eq!(run_js(r#"
function Animal(name) { this.name = name; }
Animal.prototype.speak = function() { return this.name + " speaks"; };
let a = new Animal("Dog");
console.log(a.speak());
console.log(a.hasOwnProperty("name"));
console.log(a.hasOwnProperty("speak"));
"#), &["Dog speaks", "true", "false"]);
}

#[test]
fn prototype_chain_inheritance() {
    assert_eq!(run_js(r#"
function Animal(name) { this.name = name; }
Animal.prototype.speak = function() { return "..."; };
function Dog(name) { Animal.call(this, name); }
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.speak = function() { return "Woof!"; };
let d = new Dog("Rex");
console.log(d.speak());
console.log(d.name);
console.log(d instanceof Dog);
console.log(d instanceof Animal);
"#), &["Woof!", "Rex", "true", "true"]);
}

// ===================================================================
// MISC PATTERNS
// ===================================================================

#[test]
fn tagged_template_literal() {
    assert_eq!(run_js(r#"
function upper(strings, ...values) {
    let result = "";
    strings.forEach((str, i) => {
        result += str;
        if (i < values.length) result += String(values[i]).toUpperCase();
    });
    return result;
}
let name = "world";
let num = 42;
console.log(upper`hello ${name} you are ${num}`);
"#), &["hello WORLD you are 42"]);
}

#[test]
fn structured_clone_like() {
    assert_eq!(run_js(r#"
let original = { a: 1, b: { c: 2 } };
let clone = JSON.parse(JSON.stringify(original));
clone.b.c = 99;
console.log(original.b.c);
console.log(clone.b.c);
"#), &["2", "99"]);
}

#[test]
fn optional_chaining_with_method() {
    assert_eq!(run_js(r#"
let obj = {
    foo: { bar() { return 42; } }
};
console.log(obj.foo?.bar());
console.log(obj.baz?.bar());
"#), &["42", "null"]);
}

#[test]
fn nullish_assignment_operator() {
    assert_eq!(run_js(r#"
let a = null;
a ??= 42;
console.log(a);
a ??= 99;
console.log(a);
"#), &["42", "42"]);
}

#[test]
fn logical_or_assignment() {
    assert_eq!(run_js(r#"
let a = 0;
a ||= 42;
console.log(a);
let b = "hello";
b ||= "world";
console.log(b);
"#), &["42", "hello"]);
}

#[test]
fn logical_and_assignment() {
    assert_eq!(run_js(r#"
let a = 1;
a &&= 42;
console.log(a);
let b = 0;
b &&= 42;
console.log(b);
"#), &["42", "0"]);
}
