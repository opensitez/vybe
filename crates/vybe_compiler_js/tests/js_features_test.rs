use std::cell::RefCell;
use std::rc::Rc;

/// Helper: compile + run JS, return console output lines
fn run_js(code: &str) -> Vec<String> {
    let program = vybe_parser_js::parse(code).expect("parse failed");
    let mut vm = vybe_bytecode::VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

    // Register all VSI modules + JS coercion, then override console.log to capture output
    vybe_host::register_all(&mut vm);
    vybe_compiler_js::register_js_coercion(&mut vm);
    vm.register_host_fn("vybe:console", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        out.borrow_mut().push(parts.join(" "));
        vybe_bytecode::Value::Null
    }));

    let chunks = vybe_compiler_js::Compiler::new().compile(&program).expect("compile failed");
    vm.run(chunks).expect("runtime error");
    output.borrow().clone()
}

fn run_js_one(code: &str) -> String {
    let lines = run_js(code);
    lines.into_iter().next().unwrap_or_default()
}

// ============================================================
// LITERALS & TYPES
// ============================================================

#[test]
fn test_number_literals() {
    assert_eq!(run_js_one("console.log(42)"), "42");
    assert_eq!(run_js_one("console.log(3.14)"), "3.14");
    assert_eq!(run_js_one("console.log(0)"), "0");
    assert_eq!(run_js_one("console.log(-1)"), "-1");
}

#[test]
fn test_string_literals() {
    assert_eq!(run_js_one(r#"console.log("hello")"#), "hello");
    assert_eq!(run_js_one(r#"console.log('world')"#), "world");
    assert_eq!(run_js_one(r#"console.log("it's")"#), "it's");
}

#[test]
fn test_boolean_null() {
    assert_eq!(run_js_one("console.log(true)"), "true");
    assert_eq!(run_js_one("console.log(false)"), "false");
    assert_eq!(run_js_one("console.log(null)"), "null");
}

// ============================================================
// ARITHMETIC
// ============================================================

#[test]
fn test_arithmetic() {
    assert_eq!(run_js_one("console.log(2 + 3)"), "5");
    assert_eq!(run_js_one("console.log(10 - 4)"), "6");
    assert_eq!(run_js_one("console.log(3 * 7)"), "21");
    assert_eq!(run_js_one("console.log(20 / 4)"), "5");
    assert_eq!(run_js_one("console.log(17 % 5)"), "2");
}

#[test]
fn test_unary_neg() {
    assert_eq!(run_js_one("console.log(-5)"), "-5");
    assert_eq!(run_js_one("let x = 10; console.log(-x)"), "-10");
}

#[test]
fn test_string_concat() {
    assert_eq!(run_js_one(r#"console.log("hello" + " " + "world")"#), "hello world");
    assert_eq!(run_js_one(r#"console.log("count: " + 42)"#), "count: 42");
}

// ============================================================
// VARIABLES
// ============================================================

#[test]
fn test_let_const_var() {
    assert_eq!(run_js_one("let x = 5; console.log(x)"), "5");
    assert_eq!(run_js_one("const y = 10; console.log(y)"), "10");
    assert_eq!(run_js_one("var z = 15; console.log(z)"), "15");
}

#[test]
fn test_assignment() {
    assert_eq!(run_js_one("let x = 1; x = 2; console.log(x)"), "2");
}

#[test]
fn test_compound_assignment() {
    assert_eq!(run_js_one("let x = 10; x += 5; console.log(x)"), "15");
    assert_eq!(run_js_one("let x = 10; x -= 3; console.log(x)"), "7");
    assert_eq!(run_js_one("let x = 4; x *= 3; console.log(x)"), "12");
}

// ============================================================
// COMPARISON & LOGIC
// ============================================================

#[test]
fn test_comparison() {
    assert_eq!(run_js_one("console.log(5 > 3)"), "true");
    assert_eq!(run_js_one("console.log(2 < 1)"), "false");
    assert_eq!(run_js_one("console.log(5 >= 5)"), "true");
    assert_eq!(run_js_one("console.log(4 <= 3)"), "false");
}

#[test]
fn test_strict_equality() {
    assert_eq!(run_js_one("console.log(10 === 10)"), "true");
    assert_eq!(run_js_one("console.log(10 !== 20)"), "true");
    assert_eq!(run_js_one(r#"console.log("a" === "a")"#), "true");
    assert_eq!(run_js_one(r#"console.log("a" === "b")"#), "false");
}

#[test]
fn test_logical_and_or_not() {
    assert_eq!(run_js_one("console.log(true && false)"), "false");
    assert_eq!(run_js_one("console.log(true || false)"), "true");
    assert_eq!(run_js_one("console.log(!false)"), "true");
    assert_eq!(run_js_one("console.log(!true)"), "false");
}

#[test]
fn test_short_circuit() {
    // && returns left if falsy, right if truthy
    assert_eq!(run_js_one("console.log(0 && 5)"), "0");
    assert_eq!(run_js_one("console.log(1 && 5)"), "5");
    // || returns left if truthy, right if falsy
    assert_eq!(run_js_one("console.log(0 || 5)"), "5");
    assert_eq!(run_js_one(r#"console.log("hi" || 5)"#), "hi");
}

// ============================================================
// CONTROL FLOW
// ============================================================

#[test]
fn test_if_else() {
    assert_eq!(run_js_one("if (true) { console.log('yes') } else { console.log('no') }"), "yes");
    assert_eq!(run_js_one("if (false) { console.log('yes') } else { console.log('no') }"), "no");
}

#[test]
fn test_if_without_braces() {
    assert_eq!(run_js_one("if (true) console.log('yes')"), "yes");
}

#[test]
fn test_ternary() {
    assert_eq!(run_js_one("console.log(true ? 'a' : 'b')"), "a");
    assert_eq!(run_js_one("console.log(false ? 'a' : 'b')"), "b");
}

#[test]
fn test_while_loop() {
    assert_eq!(run_js_one("let i = 0; while (i < 5) { i = i + 1; } console.log(i)"), "5");
}

#[test]
fn test_for_loop() {
    assert_eq!(run_js_one("let s = 0; for (let i = 1; i <= 10; i++) { s = s + i; } console.log(s)"), "55");
}

#[test]
fn test_do_while() {
    assert_eq!(run_js_one("let n = 1; do { n = n * 2; } while (n < 100); console.log(n)"), "128");
}

#[test]
fn test_break() {
    assert_eq!(run_js_one("let i = 0; while (true) { if (i >= 3) break; i = i + 1; } console.log(i)"), "3");
}

#[test]
fn test_continue() {
    // Sum only odd numbers 1-10
    assert_eq!(run_js_one(
        "let s = 0; for (let i = 1; i <= 10; i++) { if (i % 2 === 0) continue; s = s + i; } console.log(s)"
    ), "25"); // 1+3+5+7+9
}

#[test]
fn test_switch() {
    let code = r#"
        let x = 2;
        switch (x) {
            case 1: console.log("one"); break;
            case 2: console.log("two"); break;
            case 3: console.log("three"); break;
            default: console.log("other");
        }
    "#;
    assert_eq!(run_js_one(code), "two");
}

// ============================================================
// FUNCTIONS
// ============================================================

#[test]
fn test_function_declaration() {
    assert_eq!(run_js_one("function add(a, b) { return a + b; } console.log(add(3, 4))"), "7");
}

#[test]
fn test_recursion() {
    let code = "function fact(n) { if (n <= 1) { return 1; } return n * fact(n - 1); } console.log(fact(6))";
    assert_eq!(run_js_one(code), "720");
}

#[test]
fn test_arrow_function() {
    assert_eq!(run_js_one("let sq = (x) => x * x; console.log(sq(5))"), "25");
}

#[test]
fn test_arrow_block_body() {
    assert_eq!(run_js_one("let f = (x) => { return x + 1; }; console.log(f(9))"), "10");
}

#[test]
fn test_higher_order_function() {
    let code = r#"
        function apply(f, x) { return f(x); }
        let double = (x) => x * 2;
        console.log(apply(double, 21));
    "#;
    assert_eq!(run_js_one(code), "42");
}

#[test]
fn test_closure() {
    let code = r#"
        function makeCounter() {
            let count = 0;
            return () => { count = count + 1; return count; };
        }
        let c = makeCounter();
        console.log(c(), c(), c());
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

#[test]
fn test_compose() {
    let code = r#"
        let square = (x) => x * x;
        function compose(f, g) { return (x) => f(g(x)); }
        let ds = compose(square, (x) => x * 2);
        console.log(ds(3));
    "#;
    assert_eq!(run_js_one(code), "36");
}

// ============================================================
// OBJECTS
// ============================================================

#[test]
fn test_object_literal() {
    assert_eq!(run_js_one(r#"let o = { name: "Alice", age: 30 }; console.log(o.name, o.age)"#), "Alice 30");
}

#[test]
fn test_object_set_property() {
    assert_eq!(run_js_one(r#"let o = {}; o.x = 42; console.log(o.x)"#), "42");
}

// ============================================================
// ARRAYS
// ============================================================

#[test]
fn test_array_literal() {
    assert_eq!(run_js_one("let a = [1, 2, 3]; console.log(a)"), "1,2,3");
}

#[test]
fn test_array_length() {
    assert_eq!(run_js_one("let a = [10, 20, 30]; console.log(a.length)"), "3");
}

// ============================================================
// INCREMENT / DECREMENT
// ============================================================

#[test]
fn test_postfix_increment() {
    assert_eq!(run_js_one("let x = 5; x++; console.log(x)"), "6");
}

#[test]
fn test_prefix_increment() {
    assert_eq!(run_js_one("let x = 5; ++x; console.log(x)"), "6");
}

#[test]
fn test_postfix_returns_old() {
    assert_eq!(run_js_one("let x = 5; console.log(x++)"), "5");
}

#[test]
fn test_prefix_returns_new() {
    assert_eq!(run_js_one("let x = 5; console.log(++x)"), "6");
}

// ============================================================
// MULTIPLE ARGS TO CONSOLE.LOG
// ============================================================

#[test]
fn test_console_log_multiple() {
    assert_eq!(run_js_one("console.log(1, 2, 3)"), "1 2 3");
    assert_eq!(run_js_one(r#"console.log("x =", 42)"#), "x = 42");
}
