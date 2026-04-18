use super::helpers::run_js;

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

#[test]
fn test_01_base_class_constructor_and_method() {
    let code = r#"
        class Animal {
            constructor(name) {
                this.name = name;
            }
            speak() {
                return "I am " + this.name;
            }
        }
        let a = new Animal("dog");
        console.log(a.speak());
    "#;
    assert_eq!(run_js_one(code), "I am dog");
}

// =============================================================================
// 2. Derived extends Base -- inherited method works
// =============================================================================
#[test]
fn test_02_derived_inherits_method() {
    let code = r#"
        class Base {
            constructor(x) { this.x = x; }
            getX() { return this.x; }
        }
        class Derived extends Base {
            constructor(x) { super(x); }
        }
        let d = new Derived(42);
        console.log(d.getX());
    "#;
    assert_eq!(run_js_one(code), "42");
}

// =============================================================================
// 3. Derived overrides method
// =============================================================================
#[test]
fn test_03_derived_overrides_method() {
    let code = r#"
        class Base {
            greet() { return "base"; }
        }
        class Child extends Base {
            constructor() { super(); }
            greet() { return "child"; }
        }
        let b = new Base();
        let c = new Child();
        console.log(b.greet(), c.greet());
    "#;
    assert_eq!(run_js_one(code), "base child");
}

// =============================================================================
// 4. super() in derived constructor
// =============================================================================
#[test]
fn test_04_super_in_derived_constructor() {
    let code = r#"
        class Parent {
            constructor() { this.role = "parent"; }
        }
        class Kid extends Parent {
            constructor() {
                super();
                this.role2 = "kid";
            }
        }
        let k = new Kid();
        console.log(k.role, k.role2);
    "#;
    assert_eq!(run_js_one(code), "parent kid");
}

// =============================================================================
// 5. super() with arguments
// =============================================================================
#[test]
fn test_05_super_with_arguments() {
    let code = r#"
        class Shape {
            constructor(type, sides) {
                this.type = type;
                this.sides = sides;
            }
        }
        class Triangle extends Shape {
            constructor() {
                super("triangle", 3);
            }
        }
        let t = new Triangle();
        console.log(t.type, t.sides);
    "#;
    assert_eq!(run_js_one(code), "triangle 3");
}

// =============================================================================
// 6. super.method() calls parent version
// =============================================================================
#[test]
fn test_06_super_method_call() {
    let code = r#"
        class Base {
            greet() { return "hello"; }
        }
        class Child extends Base {
            constructor() { super(); }
            greet() { return super.greet() + " world"; }
        }
        let c = new Child();
        console.log(c.greet());
    "#;
    assert_eq!(run_js_one(code), "hello world");
}

// =============================================================================
// 7. Three levels: A -> B -> C
// =============================================================================
#[test]
fn test_07_three_level_chain() {
    let code = r#"
        class A {
            constructor() { this.a = "A"; }
            whoA() { return this.a; }
        }
        class B extends A {
            constructor() {
                super();
                this.b = "B";
            }
            whoB() { return this.b; }
        }
        class C extends B {
            constructor() {
                super();
                this.c = "C";
            }
        }
        let c = new C();
        console.log(c.whoA(), c.whoB(), c.c);
    "#;
    assert_eq!(run_js_one(code), "A B C");
}

// =============================================================================
// 8. Constructor chain through 3 levels
// =============================================================================
#[test]
fn test_08_constructor_chain_three_levels() {
    let code = r#"
        class A {
            constructor(v) { this.val = v; }
        }
        class B extends A {
            constructor(v) {
                super(v * 2);
            }
        }
        class C extends B {
            constructor(v) {
                super(v + 1);
            }
        }
        let c = new C(3);
        console.log(c.val);
    "#;
    // C(3) -> B(3+1=4) -> A(4*2=8)
    assert_eq!(run_js_one(code), "8");
}

// =============================================================================
// 9. Derived adds new method, base method still works
// =============================================================================
#[test]
fn test_09_derived_adds_method_base_still_works() {
    let code = r#"
        class Base {
            hello() { return "hi"; }
        }
        class Derived extends Base {
            constructor() { super(); }
            goodbye() { return "bye"; }
        }
        let d = new Derived();
        console.log(d.hello(), d.goodbye());
    "#;
    assert_eq!(run_js_one(code), "hi bye");
}

// =============================================================================
// 10. instanceof through chain (c instanceof A)
// =============================================================================
#[test]
fn test_10_instanceof_through_chain() {
    let code = r#"
        class A {}
        class B extends A {
            constructor() { super(); }
        }
        class C extends B {
            constructor() { super(); }
        }
        let c = new C();
        console.log(c instanceof C, c instanceof B, c instanceof A);
    "#;
    assert_eq!(run_js_one(code), "true true true");
}

// =============================================================================
// 11. Static method on base class
// =============================================================================
#[test]
fn test_11_static_method_on_base() {
    let code = r#"
        class MathHelper {
            static add(a, b) { return a + b; }
        }
        console.log(MathHelper.add(2, 3));
    "#;
    assert_eq!(run_js_one(code), "5");
}

// =============================================================================
// 12. Static method on derived class
// =============================================================================
#[test]
fn test_12_static_method_on_derived() {
    let code = r#"
        class Base {}
        class Derived extends Base {
            constructor() { super(); }
            static create(x) { return x * 10; }
        }
        console.log(Derived.create(5));
    "#;
    assert_eq!(run_js_one(code), "50");
}

// =============================================================================
// 13. Static inherited from base (Derived.staticMethod)
// =============================================================================
#[test]
fn test_13_static_inherited_from_base() {
    let code = r#"
        class Base {
            static helper() { return "from base"; }
        }
        class Derived extends Base {
            constructor() { super(); }
        }
        console.log(Derived.helper());
    "#;
    assert_eq!(run_js_one(code), "from base");
}

// =============================================================================
// 14. Method calling this.otherMethod()
// =============================================================================
#[test]
fn test_14_method_calls_this_other_method() {
    let code = r#"
        class Calc {
            double(x) { return x * 2; }
            quadruple(x) { return this.double(this.double(x)); }
        }
        let c = new Calc();
        console.log(c.quadruple(3));
    "#;
    assert_eq!(run_js_one(code), "12");
}

// =============================================================================
// 15. Override that calls super.method() then adds logic
// =============================================================================
#[test]
fn test_15_override_calls_super_then_extends() {
    let code = r#"
        class Logger {
            format(msg) { return "[LOG] " + msg; }
        }
        class TimedLogger extends Logger {
            constructor() { super(); }
            format(msg) { return super.format(msg) + " @now"; }
        }
        let tl = new TimedLogger();
        console.log(tl.format("hi"));
    "#;
    assert_eq!(run_js_one(code), "[LOG] hi @now");
}

// =============================================================================
// 16. Two derived classes from same base -- independent
// =============================================================================
#[test]
fn test_16_two_derived_classes_independent() {
    let code = r#"
        class Base {
            constructor(v) { this.v = v; }
            get() { return this.v; }
        }
        class D1 extends Base {
            constructor(v) { super(v); }
        }
        class D2 extends Base {
            constructor(v) { super(v); }
        }
        let a = new D1(10);
        let b = new D2(20);
        console.log(a.get(), b.get());
    "#;
    assert_eq!(run_js_one(code), "10 20");
}

// =============================================================================
// 17. Getter on base class
// =============================================================================
#[test]
fn test_17_getter_on_base() {
    let code = r#"
        class Circle {
            constructor(r) { this.r = r; }
            get area() { return 3.14 * this.r * this.r; }
        }
        let c = new Circle(10);
        console.log(c.area);
    "#;
    assert_eq!(run_js_one(code), "314");
}

// =============================================================================
// 18. Setter on base class
// =============================================================================
#[test]
fn test_18_setter_on_base() {
    let code = r#"
        class Container {
            constructor() { this._data = 0; }
            get data() { return this._data; }
            set data(v) { this._data = v + 1; }
        }
        let c = new Container();
        c.data = 9;
        console.log(c.data);
    "#;
    assert_eq!(run_js_one(code), "10");
}

// =============================================================================
// 19. Derived overrides getter
// =============================================================================
#[test]
fn test_19_derived_overrides_getter() {
    let code = r#"
        class Base {
            constructor() { this._x = 5; }
            get x() { return this._x; }
        }
        class Derived extends Base {
            constructor() { super(); }
            get x() { return this._x * 2; }
        }
        let b = new Base();
        let d = new Derived();
        console.log(b.x, d.x);
    "#;
    assert_eq!(run_js_one(code), "5 10");
}

// =============================================================================
// 20. No constructor in derived -- auto super()
// =============================================================================
#[test]
fn test_20_no_constructor_auto_super() {
    let code = r#"
        class Base {
            constructor() { this.ready = true; }
        }
        class Derived extends Base {}
        let d = new Derived();
        console.log(d.ready);
    "#;
    assert_eq!(run_js_one(code), "true");
}

// =============================================================================
// 21. Constructor sets this.field, method reads it
// =============================================================================
#[test]
fn test_21_constructor_sets_field_method_reads() {
    let code = r#"
        class Person {
            constructor(name, age) {
                this.name = name;
                this.age = age;
            }
            info() { return this.name + " is " + this.age; }
        }
        let p = new Person("Alice", 30);
        console.log(p.info());
    "#;
    assert_eq!(run_js_one(code), "Alice is 30");
}

// =============================================================================
// 22. Derived constructor sets own field + super() sets parent field
// =============================================================================
#[test]
fn test_22_derived_and_parent_fields() {
    let code = r#"
        class Vehicle {
            constructor(type) { this.type = type; }
        }
        class Car extends Vehicle {
            constructor(brand) {
                super("car");
                this.brand = brand;
            }
        }
        let c = new Car("Toyota");
        console.log(c.type, c.brand);
    "#;
    assert_eq!(run_js_one(code), "car Toyota");
}

// =============================================================================
// 23. Factory pattern: static create() method
// =============================================================================
#[test]
fn test_23_factory_static_create() {
    let code = r#"
        class Point {
            constructor(x, y) {
                this.x = x;
                this.y = y;
            }
            static origin() { return new Point(0, 0); }
        }
        let p = Point.origin();
        console.log(p.x, p.y);
    "#;
    assert_eq!(run_js_one(code), "0 0");
}

// =============================================================================
// 24. Method returns this (fluent API)
// =============================================================================
#[test]
#[ignore = "known limitation: new Builder().method() chain — new as direct receiver"]
fn test_24_fluent_api_returns_this() {
    let code = r#"
        class Builder {
            constructor() { this.parts = ""; }
            add(p) {
                this.parts = this.parts + p;
                return this;
            }
            build() { return this.parts; }
        }
        let result = new Builder().add("a").add("b").add("c").build();
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "abc");
}

// =============================================================================
// 25. Base class with default param values
// =============================================================================
#[test]
fn test_25_default_param_values() {
    let code = r#"
        class Config {
            constructor(host = "localhost", port = 8080) {
                this.host = host;
                this.port = port;
            }
        }
        let c1 = new Config();
        let c2 = new Config("example.com", 3000);
        console.log(c1.host, c1.port);
        console.log(c2.host, c2.port);
    "#;
    let out = run_js(code);
    assert_eq!(out[0], "localhost 8080");
    assert_eq!(out[1], "example.com 3000");
}
