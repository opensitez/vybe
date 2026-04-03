use std::cell::RefCell;
use std::rc::Rc;

fn run_js(code: &str) -> Vec<String> {
    let program = vybe_parser_js::parse(code).expect("parse failed");
    let mut vm = vybe_bytecode::VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

    // Register all VSI modules + JS coercion, then override console.log to capture output
    vybe_host::register_all(&mut vm);
    vybe_compiler_js::register_js_coercion(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[vybe_bytecode::Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        out.borrow_mut().push(parts.join(" "));
        vybe_bytecode::Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);

    let chunks = vybe_compiler_js::Compiler::new().compile(&program).expect("compile failed");
    vm.run(chunks).expect("runtime error");
    output.borrow().clone()
}

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
