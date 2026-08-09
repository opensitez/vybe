/// Class inheritance deep — super(), method resolution order, extends expression,
/// abstract class pattern, mixin inheritance, instanceof in hierarchy,
/// static inheritance, overriding methods, property shadowing.
use super::helpers::run_js;

// ── basic super() call ────────────────────────────────────────────────────────

#[test]
fn super_call_initializes_parent() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    constructor(name) { this.name = name; }
    speak() { return this.name + " speaks"; }
}
class Dog extends Animal {
    constructor(name, breed) {
        super(name);
        this.breed = breed;
    }
}
const d = new Dog("Rex", "Lab");
console.log(d.name);
console.log(d.breed);
console.log(d.speak());
"#
        ),
        vec!["Rex", "Lab", "Rex speaks"]
    );
}

#[test]
fn super_method_call() {
    assert_eq!(
        run_js(
            r#"
class Base {
    greet() { return "Base"; }
}
class Child extends Base {
    greet() { return super.greet() + "+Child"; }
}
console.log(new Child().greet());
"#
        ),
        vec!["Base+Child"]
    );
}

// ── deep inheritance chain ────────────────────────────────────────────────────

#[test]
fn three_level_inheritance_chain() {
    assert_eq!(
        run_js(
            r#"
class A {
    who() { return "A"; }
}
class B extends A {
    who() { return super.who() + "B"; }
}
class C extends B {
    who() { return super.who() + "C"; }
}
console.log(new C().who());
"#
        ),
        vec!["ABC"]
    );
}

#[test]
fn child_overrides_parent_property() {
    assert_eq!(
        run_js(
            r#"
class Shape {
    get name() { return "Shape"; }
    area() { return 0; }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    get name() { return "Circle"; }
    area() { return Math.PI * this.r * this.r; }
}
const c = new Circle(1);
console.log(c.name);
console.log(c.area().toFixed(5));
"#
        ),
        vec!["Circle", "3.14159"]
    );
}

// ── static inheritance ────────────────────────────────────────────────────────

#[test]
fn static_methods_are_inherited() {
    assert_eq!(
        run_js(
            r#"
class Base {
    type() { return "base"; }
}
class Child extends Base {
    type() { return "child"; }
}
Child.create = function() { return new Child(); };
const obj = Child.create();
console.log(obj instanceof Child);
console.log(obj.type());
"#
        ),
        vec!["true", "child"]
    );
}

#[test]
fn static_properties_inherited_by_subclass() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    static count = 0;
    static increment() { Animal.count++; }
}
class Dog extends Animal {}
Dog.increment();
console.log(Animal.count);
console.log(typeof Dog.count === "number");
"#
        ),
        vec!["1", "true"]
    );
}

// ── instanceof chain ──────────────────────────────────────────────────────────

#[test]
fn instanceof_checks_entire_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
class A {}
class B extends A {}
class C extends B {}
const c = new C();
console.log(c instanceof C);
console.log(c instanceof B);
console.log(c instanceof A);
console.log(c instanceof Object);
"#
        ),
        vec!["true", "true", "true", "true"]
    );
}

// ── extends expression ────────────────────────────────────────────────────────

#[test]
fn class_extends_function_result() {
    assert_eq!(
        run_js(
            r#"
function mixin(Base) {
    return class extends Base {
        hello() { return "hello from mixin"; }
    };
}
class Foo {}
class Bar extends mixin(Foo) {}
const b = new Bar();
console.log(b instanceof Foo);
console.log(b.hello());
"#
        ),
        vec!["true", "hello from mixin"]
    );
}

// ── mixin pattern ─────────────────────────────────────────────────────────────

#[test]
fn mixin_chain_multiple() {
    assert_eq!(
        run_js(
            r#"
const Serializable = (Base) => class extends Base {
    serialize() { return JSON.stringify(this); }
};
const Loggable = (Base) => class extends Base {
    log(msg) { return `[LOG] ${msg}`; }
};

class Model { constructor(data) { Object.assign(this, data); } }
class User extends Serializable(Loggable(Model)) {}

const u = new User({ name: "Alice", age: 30 });
console.log(u.log("hello"));
const data = JSON.parse(u.serialize());
console.log(data.name);
"#
        ),
        vec!["[LOG] hello", "Alice"]
    );
}

// ── abstract class pattern ────────────────────────────────────────────────────

#[test]
fn abstract_class_throws_when_instantiated() {
    assert_eq!(
        run_js(
            r#"
class AbstractShape {
    constructor() {
        if (this.constructor === AbstractShape) {
            throw new Error("Cannot instantiate abstract class");
        }
    }
    area() { throw new Error("Must implement area()"); }
}
class Square extends AbstractShape {
    constructor(side) { super(); this.side = side; }
    area() { return this.side ** 2; }
}
let threw = false;
try { new AbstractShape(); } catch (e) { threw = true; }
console.log(threw);
const s = new Square(4);
console.log(s.area());
"#
        ),
        vec!["true", "16"]
    );
}

// ── new.target ────────────────────────────────────────────────────────────────

#[test]
fn new_target_in_constructor() {
    assert_eq!(
        run_js(
            r#"
class Foo {
    constructor() {
        this.target = this.constructor.name;
    }
}
class Bar extends Foo {}
const f = new Foo();
const b = new Bar();
console.log(f.target);
console.log(b.target);
"#
        ),
        vec!["Foo", "Bar"]
    );
}

// ── super in static methods ───────────────────────────────────────────────────

#[test]
fn super_in_static_method() {
    assert_eq!(
        run_js(
            r#"
class A {
    static who() { return "A"; }
}
class B extends A {
    static who() { return super.who() + "B"; }
}
console.log(B.who());
"#
        ),
        vec!["AB"]
    );
}

// ── constructor return object ─────────────────────────────────────────────────

#[test]
fn constructor_returning_object_overrides_this() {
    assert_eq!(
        run_js(
            r#"
class Weird {
    constructor() {
        return { custom: true };
    }
}
const w = new Weird();
console.log(w.custom);
console.log(w instanceof Weird);
"#
        ),
        vec!["true", "false"]
    );
}

// ── property lookup order ─────────────────────────────────────────────────────

#[test]
fn own_property_shadows_prototype() {
    assert_eq!(
        run_js(
            r#"
class Base {
    get value() { return "prototype"; }
}
class Child extends Base {
    constructor() {
        super();
        Object.defineProperty(this, "value", { value: "own", writable: true, configurable: true, enumerable: true });
    }
}
const c = new Child();
console.log(c.value);
"#
        ),
        vec!["own"]
    );
}

#[test]
fn derived_constructor_runs_field_initializers_before_body() {
    assert_eq!(
        run_js(
            r#"
const order = [];
class Base {
    constructor() { order.push("base"); }
}
class Derived extends Base {
    initialized = (order.push("field"), 1);
    constructor() {
        super();
        order.push("constructor");
    }
}
new Derived();
console.log(order.join("|"));
"#
        ),
        vec!["base|field|constructor"]
    );
}

#[test]
fn static_super_chain_calls_base_methods() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    static who() { return "Animal"; }
}
class Mammal extends Animal {
    static who() { return super.who() + ":Mammal"; }
}
class Dog extends Mammal {
    static who() { return super.who() + ":Dog"; }
}
console.log(Dog.who());
"#
        ),
        vec!["Animal:Mammal:Dog"]
    );
}

#[test]
fn static_super_reads_base_static_property() {
    assert_eq!(
        run_js(
            r#"
class Base {
    static marker = "base";
}
class Child extends Base {
    static marker = "child";
    static getBaseMarker() {
        return super.marker;
    }
}
console.log(Child.marker);
console.log(Child.getBaseMarker());
console.log(Base.marker);
"#
        ),
        vec!["child", "base", "base"]
    );
}

#[test]
fn derived_constructor_without_super_throws_reference_error() {
    assert_eq!(
        run_js(
            r#"
class Base {}
class Broken extends Base {
    constructor() {
        this.x = 1;
    }
}

let threw = false;
try {
    new Broken();
} catch (e) {
    threw = e instanceof ReferenceError;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn derived_instance_field_shadows_base_getter_property() {
    assert_eq!(
        run_js(
            r#"
class Base {
    constructor() {
        this.mode = "base-mode";
    }
}
class Child extends Base {
    constructor() {
        super();
        this.mode = "field-mode";
    }
}
const child = new Child();
console.log(`${child.mode}|${Object.hasOwn(child, "mode")}|${child.mode === "field-mode"}`);
"#
        ),
        vec!["field-mode|true|true"]
    );
}

#[test]
fn test_static_super_method_preserves_dynamic_this_receiver() {
    assert_eq!(
        run_js(
            r#"
class Base {
    static getName() {
        return this.name;
    }
}
class Derived extends Base {
    static getName() {
        return super.getName() + "Suffix";
    }
}
console.log(Derived.getName());
"#
        ),
        vec!["DerivedSuffix"]
    );
}
