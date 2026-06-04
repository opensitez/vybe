/// Class inheritance edge cases — super in static, extends expression, new.target
use super::helpers::run_js;

#[test]
fn super_in_constructor_chain() {
    assert_eq!(
        run_js(
            r#"
class A {
    constructor(x) { this.x = x; }
}
class B extends A {
    constructor(x, y) {
        super(x);
        this.y = y;
    }
}
class C extends B {
    constructor(x, y, z) {
        super(x, y);
        this.z = z;
    }
}
const c = new C(1, 2, 3);
console.log(c.x);
console.log(c.y);
console.log(c.z);
"#
        ),
        vec!["1", "2", "3"]
    );
}

#[test]
fn super_method_in_static_context() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    static describe() { return "Animal"; }
}
class Dog extends Animal {
    static describe() { return super.describe() + "/Dog"; }
}
console.log(Dog.describe());
"#
        ),
        vec!["Animal/Dog"]
    );
}

#[test]
fn new_target_is_subclass_in_super_constructor() {
    assert_eq!(
        run_js(
            r#"
class Base {
    constructor() {
        this.constructedAs = this.constructor.name;
    }
}
class Derived extends Base {}
const b = new Base();
const d = new Derived();
console.log(b.constructedAs);
console.log(d.constructedAs);
"#
        ),
        vec!["Base", "Derived"]
    );
}

#[test]
fn extends_with_expression() {
    assert_eq!(
        run_js(
            r#"
function makeBase(msg) {
    return class {
        greet() { return msg; }
    };
}
const Base = makeBase("hello from factory");
class Derived extends Base {}
const d = new Derived();
console.log(d.greet());
"#
        ),
        vec!["hello from factory"]
    );
}

#[test]
fn subclass_overrides_getter() {
    assert_eq!(
        run_js(
            r#"
class Shape {
    get area() { return 0; }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    get area() { return Math.PI * this.r * this.r; }
}
const c = new Circle(1);
console.log(c.area.toFixed(5));
"#
        ),
        vec!["3.14159"]
    );
}

#[test]
fn parent_method_accessible_via_super() {
    assert_eq!(
        run_js(
            r#"
class Logger {
    log(msg) { return "[LOG] " + msg; }
}
class PrefixLogger extends Logger {
    constructor(prefix) {
        super();
        this.prefix = prefix;
    }
    log(msg) {
        return super.log(this.prefix + ": " + msg);
    }
}
const logger = new PrefixLogger("App");
console.log(logger.log("started"));
"#
        ),
        vec!["[LOG] App: started"]
    );
}

#[test]
fn class_in_expression_position() {
    assert_eq!(
        run_js(
            r#"
const Greeter = class NamedGreeter {
    greet(name) { return "Hello " + name; }
};
const g = new Greeter();
console.log(g.greet("World"));
console.log(typeof g.greet);
"#
        ),
        vec!["Hello World", "function"]
    );
}

#[test]
fn instanceof_in_prototype_chain() {
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
const b = new B();
console.log(b instanceof C); // false — b is not a C
"#
        ),
        vec!["true", "true", "true", "true", "false"]
    );
}

#[test]
fn subclass_calls_super_method_with_this() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    constructor() { this.count = 0; }
    increment() { this.count++; return this; }
}
class BoundedCounter extends Counter {
    constructor(max) {
        super();
        this.max = max;
    }
    increment() {
        if (this.count < this.max) super.increment();
        return this;
    }
}
const bc = new BoundedCounter(3);
bc.increment().increment().increment().increment().increment();
console.log(bc.count);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn inherits_prototype_methods() {
    assert_eq!(
        run_js(
            r#"
class EventEmitter {
    constructor() { this._handlers = {}; }
    on(event, fn) {
        (this._handlers[event] = this._handlers[event] || []).push(fn);
    }
    emit(event, ...args) {
        (this._handlers[event] || []).forEach(fn => fn(...args));
    }
}
class Button extends EventEmitter {
    click() { this.emit("click", this); }
}
const btn = new Button();
const log = [];
btn.on("click", () => log.push("clicked"));
btn.click();
btn.click();
console.log(log.join(","));
"#
        ),
        vec!["clicked,clicked"]
    );
}
