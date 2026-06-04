use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Classes — ES2015 through ES2024
// ═══════════════════════════════════════════════════════════

#[test]
fn class_basic() {
    let out = run_js(
        r#"
class Animal {
    constructor(name) {
        this.name = name;
    }
    speak() {
        return this.name + " makes a noise";
    }
}
const a = new Animal("Dog");
console.log(a.speak());
"#,
    );
    assert_eq!(out, vec!["Dog makes a noise"]);
}

#[test]
fn class_inheritance() {
    let out = run_js(
        r#"
class Shape {
    constructor(color) {
        this.color = color;
    }
    describe() {
        return "A " + this.color + " shape";
    }
}
class Circle extends Shape {
    constructor(color, radius) {
        super(color);
        this.radius = radius;
    }
    area() {
        return 3.14159 * this.radius * this.radius;
    }
}
const c = new Circle("red", 5);
console.log(c.describe());
console.log(c.area());
"#,
    );
    assert_eq!(out, vec!["A red shape", "78.53975"]);
}

#[test]
fn class_super_method() {
    let out = run_js(
        r#"
class Base {
    greet() { return "Hello"; }
}
class Derived extends Base {
    greet() { return super.greet() + " World"; }
}
const d = new Derived();
console.log(d.greet());
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn class_static_method() {
    let out = run_js(
        r#"
class MathUtils {
    static square(x) { return x * x; }
    static cube(x) { return x * x * x; }
}
console.log(MathUtils.square(4));
console.log(MathUtils.cube(3));
"#,
    );
    assert_eq!(out, vec!["16", "27"]);
}

#[test]
fn class_static_property() {
    let out = run_js(
        r#"
class Config {
    static version = "1.0.0";
    static debug = false;
}
console.log(Config.version);
console.log(Config.debug);
"#,
    );
    assert_eq!(out, vec!["1.0.0", "false"]);
}

#[test]
fn class_getter() {
    let out = run_js(
        r#"
class Rectangle {
    constructor(w, h) {
        this.width = w;
        this.height = h;
    }
    get area() {
        return this.width * this.height;
    }
}
const r = new Rectangle(5, 3);
console.log(r.area);
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn class_setter() {
    let out = run_js(
        r#"
class Temperature {
    constructor(celsius) {
        this._celsius = celsius;
    }
    get fahrenheit() {
        return this._celsius * 9 / 5 + 32;
    }
    set fahrenheit(f) {
        this._celsius = (f - 32) * 5 / 9;
    }
}
const t = new Temperature(0);
console.log(t.fahrenheit);
t.fahrenheit = 212;
console.log(t._celsius);
"#,
    );
    assert_eq!(out, vec!["32", "100"]);
}

#[test]
fn class_computed_method_name() {
    let out = run_js(
        r#"
const methodName = "greet";
class Greeter {
    [methodName]() {
        return "Hello!";
    }
}
const g = new Greeter();
console.log(g.greet());
"#,
    );
    assert_eq!(out, vec!["Hello!"]);
}

#[test]
fn class_expression() {
    let out = run_js(
        r#"
const MyClass = class {
    constructor(val) { this.val = val; }
    getVal() { return this.val; }
};
const obj = new MyClass(42);
console.log(obj.getVal());
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn class_expression_named() {
    let out = run_js(
        r#"
const Foo = class Bar {
    name() { return "Bar"; }
};
const f = new Foo();
console.log(f.name());
"#,
    );
    assert_eq!(out, vec!["Bar"]);
}

#[test]
fn class_private_field() {
    let out = run_js(
        r#"
class Counter {
    #count = 0;
    increment() { this.#count++; }
    get value() { return this.#count; }
}
const c = new Counter();
c.increment();
c.increment();
c.increment();
console.log(c.value);
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn class_private_method() {
    let out = run_js(
        r#"
class Validator {
    #validate(input) {
        return input.length > 0;
    }
    check(input) {
        return this.#validate(input);
    }
}
const v = new Validator();
console.log(v.check("hello"));
console.log(v.check(""));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn class_multi_level_inheritance() {
    let out = run_js(
        r#"
class A {
    whoami() { return "A"; }
}
class B extends A {
    whoami() { return "B->" + super.whoami(); }
}
class C extends B {
    whoami() { return "C->" + super.whoami(); }
}
const c = new C();
console.log(c.whoami());
"#,
    );
    assert_eq!(out, vec!["C->B->A"]);
}

#[test]
fn class_instanceof() {
    let out = run_js(
        r#"
class Animal {}
class Dog extends Animal {}
const d = new Dog();
console.log(d instanceof Dog);
console.log(d instanceof Animal);
"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn class_method_chaining() {
    let out = run_js(
        r#"
class Builder {
    constructor() { this.parts = []; }
    add(part) { this.parts.push(part); return this; }
    build() { return this.parts.join(", "); }
}
const result = new Builder().add("a").add("b").add("c").build();
console.log(result);
"#,
    );
    assert_eq!(out, vec!["a, b, c"]);
}

#[test]
fn class_property_initializer() {
    let out = run_js(
        r#"
class Defaults {
    name = "unnamed";
    count = 0;
    items = [];

    describe() {
        return this.name + ":" + this.count;
    }
}
const d = new Defaults();
console.log(d.describe());
"#,
    );
    assert_eq!(out, vec!["unnamed:0"]);
}

#[test]
fn class_extends_expression() {
    let out = run_js(
        r#"
function getBase() {
    return class {
        greet() { return "base"; }
    };
}
class Derived extends getBase() {
    greet() { return super.greet() + "+derived"; }
}
const d = new Derived();
console.log(d.greet());
"#,
    );
    assert_eq!(out, vec!["base+derived"]);
}

#[test]
fn class_to_string_override() {
    let out = run_js(
        r#"
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
    toString() { return "(" + this.x + ", " + this.y + ")"; }
}
const p = new Point(3, 4);
console.log(p.toString());
"#,
    );
    assert_eq!(out, vec!["(3, 4)"]);
}
