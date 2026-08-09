/// Class field declarations — public/private instance fields, static fields,
/// field initializers, class auto-accessors, private in-check (#field in obj).
use super::helpers::run_js;

// ── public instance fields ────────────────────────────────────────────────────

#[test]
fn public_field_has_default_undefined() {
    assert_eq!(
        run_js(
            r#"
class Foo {
    x;
    y = 42;
}
const f = new Foo();
console.log(f.x);
console.log(f.y);
"#
        ),
        vec!["undefined", "42"]
    );
}

#[test]
fn public_field_initialized_before_constructor_body() {
    assert_eq!(
        run_js(
            r#"
class Child {
    compute() { return 5; }
    constructor() { this.value = this.compute(); }
}
const c = new Child();
console.log(c.value);
"#
        ),
        vec!["5"]
    );
}

#[test]
fn public_fields_are_per_instance() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    count = 0;
    increment() { this.count++; }
}
const a = new Counter();
const b = new Counter();
a.increment(); a.increment();
b.increment();
console.log(a.count);
console.log(b.count);
"#
        ),
        vec!["2", "1"]
    );
}

// ── private instance fields ───────────────────────────────────────────────────

#[test]
fn private_field_not_accessible_outside() {
    assert_eq!(
        run_js(
            r#"
class Secret {
    #value = 42;
    get() { return this.#value; }
}
const s = new Secret();
console.log(s.get());
let threw = false;
try { s.#value; } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["42", "true"]
    );
}

#[test]
fn private_field_in_operator() {
    assert_eq!(
        run_js(
            r#"
class Tagged {
    #tag = true;
    static isTagged(obj) { return #tag in obj; }
}
const t = new Tagged();
console.log(Tagged.isTagged(t));
console.log(Tagged.isTagged({}));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn private_field_brand_check() {
    assert_eq!(
        run_js(
            r#"
class Circle {
    #radius;
    constructor(r) { this.#radius = r; }
    static isCircle(obj) { return #radius in obj; }
    area() { return Math.PI * this.#radius ** 2; }
}
const c = new Circle(5);
console.log(Circle.isCircle(c));
console.log(Circle.isCircle({}));
"#
        ),
        vec!["true", "false"]
    );
}

// ── static fields ─────────────────────────────────────────────────────────────

#[test]
fn static_field_shared_on_class() {
    assert_eq!(
        run_js(
            r#"
class Config {
    static defaultTimeout = 5000;
    static VERSION = "1.0.0";
}
console.log(Config.defaultTimeout);
console.log(Config.VERSION);
// Not on instances
const c = new Config();
console.log(c.defaultTimeout === undefined);
"#
        ),
        vec!["5000", "1.0.0", "true"]
    );
}

#[test]
fn static_private_field() {
    assert_eq!(
        run_js(
            r#"
class IdGenerator {
    static #nextId = 0;
    static generate() { return ++IdGenerator.#nextId; }
}
console.log(IdGenerator.generate());
console.log(IdGenerator.generate());
console.log(IdGenerator.generate());
"#
        ),
        vec!["1", "2", "3"]
    );
}

// ── static initialization blocks ─────────────────────────────────────────────

#[test]
fn static_block_initializes_complex_state() {
    assert_eq!(
        run_js(
            r#"
class Config {
    static #data = new Map([["a", 1], ["b", 2]]);
    static get(key) { return Config.#data.get(key); }
}
console.log(Config.get("a"));
console.log(Config.get("b"));
"#
        ),
        vec!["1", "2"]
    );
}

// ── auto-accessors ────────────────────────────────────────────────────────────

#[test]
fn auto_accessor_generates_getter_and_setter() {
    assert_eq!(
        run_js(
            r#"
class Temp {
    accessor celsius = 0;
    get fahrenheit() { return this.celsius * 1.8 + 32; }
}
const t = new Temp();
t.celsius = 100;
console.log(t.fahrenheit);
console.log(t.celsius);
"#
        ),
        vec!["212", "100"]
    );
}

// ── field initializer order ────────────────────────────────────────────────────

#[test]
fn field_initializer_order_within_class() {
    assert_eq!(
        run_js(
            r#"
const log = [];
class Ordered {
    a = log.push("a") && 1;
    b = log.push("b") && 2;
    c = log.push("c") && 3;
    constructor() { log.push("ctor"); }
}
new Ordered();
console.log(log.join(","));
"#
        ),
        vec!["a,b,c,ctor"]
    );
}

// ── computed field names ──────────────────────────────────────────────────────

#[test]
fn computed_class_field_name() {
    assert_eq!(
        run_js(
            r#"
const fieldName = "dynamic";
class Dyn {
    constructor() { this[fieldName] = 42; }
}
const d = new Dyn();
console.log(d.dynamic);
"#
        ),
        vec!["42"]
    );
}

// ── private methods ───────────────────────────────────────────────────────────

#[test]
fn private_method_only_callable_inside() {
    assert_eq!(
        run_js(
            r#"
class Processor {
    #transform(x) { return x * 2; }
    process(x) { return this.#transform(x); }
}
const p = new Processor();
console.log(p.process(21));
let threw = false;
try { p.#transform(1); } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["42", "true"]
    );
}

// ── field vs method performance pattern ──────────────────────────────────────

#[test]
fn arrow_field_vs_method_binding() {
    assert_eq!(
        run_js(
            r#"
class Handler {
    name = "handler";
    // Arrow as field binds 'this' permanently
    arrowMethod = () => this.name;
    // Regular method — 'this' depends on call site
    regularMethod() { return this.name; }
}
const h = new Handler();
const { arrowMethod, regularMethod } = h;
console.log(arrowMethod()); // works — bound
let threw = false;
try {
    const r = regularMethod(); // might throw or return undefined
} catch { threw = true; }
// Either throws (strict mode) or returns undefined (sloppy)
console.log(typeof arrowMethod());
"#
        ),
        vec!["handler", "string"]
    );
}

#[test]
fn field_initializers_run_in_order_and_access_this() {
    assert_eq!(
        run_js(
            r#"
class Sequence {
    first = 1;
    second = this.first + 1;
}
const s = new Sequence();
console.log(s.first);
console.log(s.second);
"#
        ),
        vec!["1", "2"]
    );
}

#[test]
fn class_public_field_is_own_enumerable_data_property() {
    assert_eq!(
        run_js(
            r#"
class Config {
    enabled = true;
}
const c = new Config();
const desc = Object.getOwnPropertyDescriptor(c, "enabled");
console.log(desc !== undefined);
console.log(desc.enumerable);
console.log(desc.writable);
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn derived_instance_field_overrides_base_field() {
    assert_eq!(
        run_js(
            r#"
class Base {
    x = 1;
}

class Derived extends Base {
    x = 2;
}

const d = new Derived();
console.log(d.x);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn inherited_instance_field_and_subclass_field_combine() {
    assert_eq!(
        run_js(
            r#"
class Base {
    base = "base";
}

class Child extends Base {
    child = "child";
}

const c = new Child();
console.log(c.base);
console.log(c.child);
"#
        ),
        vec!["base", "child"]
    );
}

#[test]
fn class_static_fields_do_not_inherit_instance_shape() {
    assert_eq!(
        run_js(
            r#"
class Parent {
    static role = "parent";
    instanceRole = "instance-parent";
}

class Child extends Parent {
    static role = "child";
}

const p = new Parent();
const c = new Child();
console.log(p.instanceRole);
console.log(Parent.role);
console.log(Child.role);
console.log(c.instanceRole);
"#
        ),
        vec!["instance-parent", "parent", "child", "instance-parent"]
    );
}

#[test]
fn test_derived_field_initializers_run_after_super_returns() {
    assert_eq!(
        run_js(
            r#"
let baseSeen = "uninitialized";
class Base {
    constructor() {
        baseSeen = String(this.derivedField);
    }
}
class Derived extends Base {
    derivedField = "initialized";
}
const d = new Derived();
console.log(`${baseSeen}|${d.derivedField}`);
"#
        ),
        vec!["undefined|initialized"]
    );
}
