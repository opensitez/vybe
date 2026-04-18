use super::helpers::run_js;

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

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
