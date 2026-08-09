use super::helpers::run_js;

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

// ============================================================
// BASIC CLASS
// ============================================================

#[test]
fn test_class_constructor() {
    let code = r#"
        class Point {
            constructor(x, y) {
                this.x = x;
                this.y = y;
            }
        }
        let p = new Point(3, 4);
        console.log(p.x, p.y);
    "#;
    assert_eq!(run_js_one(code), "3 4");
}

#[test]
fn test_class_method() {
    let code = r#"
        class Point {
            constructor(x, y) {
                this.x = x;
                this.y = y;
            }
            sum() {
                return this.x + this.y;
            }
        }
        let p = new Point(3, 4);
        console.log(p.sum());
    "#;
    assert_eq!(run_js_one(code), "7");
}

#[test]
fn test_class_multiple_methods() {
    let code = r#"
        class Counter {
            constructor(start) {
                this.count = start;
            }
            increment() {
                this.count = this.count + 1;
            }
            get() {
                return this.count;
            }
        }
        let c = new Counter(0);
        c.increment();
        c.increment();
        c.increment();
        console.log(c.get());
    "#;
    assert_eq!(run_js_one(code), "3");
}

#[test]
fn test_class_default_constructor() {
    let code = r#"
        class Empty {}
        let e = new Empty();
        console.log(typeof e);
    "#;
    // typeof returns "object" for our objects
    assert_eq!(run_js_one(code), "object");
}

#[test]
fn test_class_with_string_method() {
    let code = r#"
        class Greeter {
            constructor(name) {
                this.name = name;
            }
            greet() {
                return "Hello, " + this.name + "!";
            }
        }
        let g = new Greeter("World");
        console.log(g.greet());
    "#;
    assert_eq!(run_js_one(code), "Hello, World!");
}

#[test]
fn test_multiple_instances() {
    let code = r#"
        class Box {
            constructor(w, h) {
                this.w = w;
                this.h = h;
            }
            area() {
                return this.w * this.h;
            }
        }
        let a = new Box(3, 4);
        let b = new Box(5, 6);
        console.log(a.area(), b.area());
    "#;
    assert_eq!(run_js_one(code), "12 30");
}

#[test]
fn test_class_property_set() {
    let code = r#"
        class Dog {
            constructor(name) {
                this.name = name;
                this.tricks = 0;
            }
            learn() {
                this.tricks = this.tricks + 1;
            }
        }
        let d = new Dog("Rex");
        d.learn();
        d.learn();
        console.log(d.name, d.tricks);
    "#;
    assert_eq!(run_js_one(code), "Rex 2");
}

// ============================================================
// CLASS + CLOSURES
// ============================================================

#[test]
fn test_class_with_closure() {
    let code = r#"
        class Adder {
            constructor(base) {
                this.base = base;
            }
            add(x) {
                return this.base + x;
            }
        }
        let a = new Adder(100);
        console.log(a.add(42));
    "#;
    assert_eq!(run_js_one(code), "142");
}

// ============================================================
// CLASS + STDLIB
// ============================================================

#[test]
fn test_class_with_array() {
    let code = r#"
        class Stack {
            constructor() {
                this.items = [];
            }
            push(item) {
                this.items.push(item);
            }
            size() {
                return this.items.length;
            }
        }
        let s = new Stack();
        s.push(1);
        s.push(2);
        s.push(3);
        console.log(s.items.length, s.size());
    "#;
    assert_eq!(run_js_one(code), "3 3");
}

#[test]
fn test_class_getter_setter() {
    let code = r#"
        class Temperature {
            constructor(celsius) {
                this.celsius = celsius;
            }
            get fahrenheit() {
                return this.celsius * 9 / 5 + 32;
            }
            set fahrenheit(value) {
                this.celsius = (value - 32) * 5 / 9;
            }
        }
        const t = new Temperature(0);
        const f = t.fahrenheit;
        t.fahrenheit = 212;
        console.log(f, t.celsius);
    "#;
    assert_eq!(run_js_one(code), "32 100");
}

#[test]
fn test_class_static_methods() {
    let code = r#"
        class MathTools {
            static add(a, b) {
                return a + b;
            }
            static triple(v) {
                return MathTools.add(v, v + v);
            }
        }
        const total = MathTools.add(4, 5);
        const triple = MathTools.triple(7);
        console.log(total, triple);
    "#;
    assert_eq!(run_js_one(code), "9 21");
}

#[test]
fn test_class_inheritance_super_call() {
    let code = r#"
        class Base {
            constructor(name) {
                this.kind = "base";
                this.name = name;
            }
            describe() {
                return this.kind + ":" + this.name;
            }
        }
        class Derived extends Base {
            constructor(name, role) {
                super(name);
                this.role = role;
            }
            describe() {
                return super.describe() + ":" + this.role;
            }
        }
        const d = new Derived("alpha", "admin");
        console.log(d.describe());
    "#;
    assert_eq!(run_js_one(code), "base:alpha:admin");
}

#[test]
fn test_class_constructor_explicit_object_return() {
    let code = r#"
        class Custom {
            constructor() {
                return { custom: true };
            }
        }
        const obj = new Custom();
        console.log(obj.custom);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_class_constructor_primitive_return_ignored() {
    let code = r#"
        class Prim {
            constructor() {
                this.val = 42;
                return 123;
            }
        }
        const obj = new Prim();
        console.log(obj.val);
    "#;
    assert_eq!(run_js_one(code), "42");
}

#[test]
fn test_class_constructor_call_without_new_throws_typeerror() {
    let code = r#"
        class C {}
        try {
            C();
        } catch (e) {
            console.log(e.name);
        }
    "#;
    assert_eq!(run_js_one(code), "TypeError");
}
