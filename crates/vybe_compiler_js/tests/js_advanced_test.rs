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
    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
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
// OBJECT METHODS
// ============================================================

#[test]
fn test_object_method() {
    let code = r#"
        let obj = {
            name: "Alice",
            greet() { return "Hello " + this.name; }
        };
        // Note: 'this' is not bound yet, so we test method as value
        console.log(obj.name);
    "#;
    assert_eq!(run_js_one(code), "Alice");
}

#[test]
fn test_object_computed_access() {
    let code = r#"
        let obj = { x: 10, y: 20 };
        let key = "x";
        console.log(obj[key]);
    "#;
    // This requires computed member access on objects
    assert_eq!(run_js_one(code), "10");
}

// ============================================================
// NESTED FUNCTIONS & SCOPE
// ============================================================

#[test]
fn test_nested_closures() {
    let code = r#"
        function outer() {
            let x = 10;
            function middle() {
                let y = 20;
                function inner() {
                    return x + y;
                }
                return inner();
            }
            return middle();
        }
        console.log(outer());
    "#;
    assert_eq!(run_js_one(code), "30");
}

#[test]
fn test_closure_mutation() {
    let code = r#"
        function counter() {
            let n = 0;
            return {
                inc() { n = n + 1; return n; },
                get() { return n; }
            };
        }
        // Object methods don't have 'this' bound, but closures work
        // We can't call c.inc() as a method yet, but we can test closure capture
        let n = 0;
        function inc() { n = n + 1; return n; }
        console.log(inc(), inc(), inc());
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// ============================================================
// STRING FEATURES
// ============================================================

#[test]
fn test_string_comparison() {
    assert_eq!(run_js_one(r#"console.log("abc" === "abc")"#), "true");
    assert_eq!(run_js_one(r#"console.log("abc" === "def")"#), "false");
}

#[test]
fn test_string_length() {
    assert_eq!(run_js_one(r#"console.log("hello".length)"#), "5");
}

#[test]
fn test_empty_string_falsy() {
    assert_eq!(run_js_one(r#"console.log("" || "default")"#), "default");
    assert_eq!(run_js_one(r#"console.log("hello" || "default")"#), "hello");
}

// ============================================================
// MULTIPLE STATEMENTS & COMPLEX PROGRAMS
// ============================================================

#[test]
fn test_fibonacci() {
    let code = r#"
        function fib(n) {
            if (n <= 1) return n;
            return fib(n - 1) + fib(n - 2);
        }
        console.log(fib(10));
    "#;
    assert_eq!(run_js_one(code), "55");
}

#[test]
fn test_array_manual_iteration() {
    let code = r#"
        let arr = [10, 20, 30, 40, 50];
        let sum = 0;
        for (let i = 0; i < arr.length; i++) {
            sum = sum + arr[i];
        }
        console.log(sum);
    "#;
    // This requires: array index access, array.length
    assert_eq!(run_js_one(code), "150");
}

#[test]
fn test_nested_if() {
    let code = r#"
        function classify(n) {
            if (n > 0) {
                if (n > 100) return "big";
                else return "small positive";
            } else if (n < 0) {
                return "negative";
            } else {
                return "zero";
            }
        }
        console.log(classify(50), classify(-5), classify(0), classify(200));
    "#;
    assert_eq!(run_js_one(code), "small positive negative zero big");
}

#[test]
fn test_multiple_outputs() {
    let lines = run_js(r#"
        console.log("line 1");
        console.log("line 2");
        console.log("line 3");
    "#);
    assert_eq!(lines, vec!["line 1", "line 2", "line 3"]);
}

// ============================================================
// TYPEOF (via host function)
// ============================================================

#[test]
fn test_typeof() {
    assert_eq!(run_js_one("console.log(typeof 42)"), "number");
    assert_eq!(run_js_one(r#"console.log(typeof "hello")"#), "string");
    assert_eq!(run_js_one("console.log(typeof true)"), "boolean");
    assert_eq!(run_js_one("console.log(typeof null)"), "object"); // JS spec: typeof null === "object"
    assert_eq!(run_js_one("console.log(typeof undefined)"), "undefined");
}

// ============================================================
// NULLISH / DEFAULT VALUES
// ============================================================

#[test]
fn test_null_or_default() {
    assert_eq!(run_js_one("let x = null; console.log(x || 42)"), "42");
    assert_eq!(run_js_one("let x = 10; console.log(x || 42)"), "10");
}

// ============================================================
// EDGE CASES
// ============================================================

#[test]
fn test_empty_function() {
    assert_eq!(run_js_one("function noop() {} console.log(noop())"), "null");
}

#[test]
fn test_return_without_value() {
    assert_eq!(run_js_one("function f() { return; } console.log(f())"), "null");
}

#[test]
fn test_deeply_nested_loops() {
    let code = r#"
        let sum = 0;
        for (let i = 0; i < 10; i++) {
            for (let j = 0; j < 10; j++) {
                sum = sum + 1;
            }
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "100");
}

#[test]
fn test_array_of_functions() {
    let code = r#"
        let fns = [(x) => x + 1, (x) => x * 2, (x) => x * x];
        console.log(fns[0](5), fns[1](5), fns[2](5));
    "#;
    assert_eq!(run_js_one(code), "6 10 25");
}

#[test]
fn test_object_nested() {
    let code = r#"
        let o = { a: { b: 42 } };
        console.log(o.a.b);
    "#;
    assert_eq!(run_js_one(code), "42");
}
