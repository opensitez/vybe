use std::cell::RefCell;
use std::rc::Rc;

/// Helper: compile + run JS, return console output lines
fn run_js(code: &str) -> Vec<String> {
    let program = vybe_parser_js::parse(code).expect("parse failed");
    let mut vm = vybe_bytecode::VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

    vybe_host::register_all(&mut vm);
    vybe_compiler_js::register_js_coercion(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
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
// Exponentiation operator **
// ============================================================

#[test]
fn test_exp_operator() {
    assert_eq!(run_js_one("console.log(2 ** 3)"), "8");
}

#[test]
fn test_exp_operator_fractional() {
    assert_eq!(run_js_one("console.log(9 ** 0.5)"), "3");
}

#[test]
fn test_exp_assign() {
    assert_eq!(run_js_one("let x = 2; x **= 10; console.log(x)"), "1024");
}

// ============================================================
// instanceof operator
// ============================================================

#[test]
fn test_instanceof_class() {
    let code = r#"
        class Foo {}
        let f = new Foo();
        console.log(f instanceof Foo);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_instanceof_wrong_class() {
    let code = r#"
        class Foo {}
        class Bar {}
        let f = new Foo();
        console.log(f instanceof Bar);
    "#;
    assert_eq!(run_js_one(code), "false");
}

// ============================================================
// in operator
// ============================================================

#[test]
fn test_in_operator_found() {
    let code = r#"
        let obj = { name: "test", age: 25 };
        console.log("name" in obj);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_in_operator_not_found() {
    let code = r#"
        let obj = { name: "test" };
        console.log("age" in obj);
    "#;
    assert_eq!(run_js_one(code), "false");
}

// ============================================================
// delete operator
// ============================================================

#[test]
fn test_delete_property() {
    let code = r#"
        let obj = { a: 1, b: 2 };
        delete obj.a;
        console.log("a" in obj);
    "#;
    assert_eq!(run_js_one(code), "false");
}

#[test]
fn test_delete_keeps_other_props() {
    let code = r#"
        let obj = { a: 1, b: 2 };
        delete obj.a;
        console.log(obj.b);
    "#;
    assert_eq!(run_js_one(code), "2");
}

// ============================================================
// Labeled break / continue
// ============================================================

#[test]
fn test_labeled_break() {
    let code = r#"
        let result = 0;
        outer: for (let i = 0; i < 3; i++) {
            for (let j = 0; j < 3; j++) {
                if (i === 1 && j === 1) break outer;
                result++;
            }
        }
        console.log(result);
    "#;
    // i=0: j=0,1,2 → 3; i=1: j=0 → 1, j=1 → break outer. Total: 4
    assert_eq!(run_js_one(code), "4");
}

#[test]
fn test_labeled_continue() {
    let code = r#"
        let result = 0;
        outer: for (let i = 0; i < 3; i++) {
            for (let j = 0; j < 3; j++) {
                if (j === 1) continue outer;
                result++;
            }
        }
        console.log(result);
    "#;
    // Each outer iteration: j=0 → count, j=1 → continue outer. So 3 iterations, 1 count each = 3
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// dyn_ne (NaN handling)
// ============================================================

#[test]
fn test_dyn_ne_nan() {
    assert_eq!(run_js_one("console.log(NaN != NaN)"), "true");
}

#[test]
fn test_dyn_ne_strings() {
    assert_eq!(run_js_one(r#"console.log("a" != "b")"#), "true");
    assert_eq!(run_js_one(r#"console.log("a" != "a")"#), "false");
}

// ============================================================
// dyn_le / dyn_ge with strings
// ============================================================

#[test]
fn test_dyn_le_strings() {
    assert_eq!(run_js_one(r#"console.log("a" <= "b")"#), "true");
    assert_eq!(run_js_one(r#"console.log("b" <= "a")"#), "false");
    assert_eq!(run_js_one(r#"console.log("a" <= "a")"#), "true");
}

#[test]
fn test_dyn_ge_strings() {
    assert_eq!(run_js_one(r#"console.log("b" >= "a")"#), "true");
    assert_eq!(run_js_one(r#"console.log("a" >= "b")"#), "false");
    assert_eq!(run_js_one(r#"console.log("a" >= "a")"#), "true");
}

// ============================================================
// Array.some / every / findIndex
// ============================================================

#[test]
fn test_array_some_true() {
    let code = r#"
        let arr = [1, 2, 3, 4];
        console.log(arr.some(x => x > 3));
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_array_some_false() {
    let code = r#"
        let arr = [1, 2, 3];
        console.log(arr.some(x => x > 5));
    "#;
    assert_eq!(run_js_one(code), "false");
}

#[test]
fn test_array_every_true() {
    let code = r#"
        let arr = [2, 4, 6];
        console.log(arr.every(x => x % 2 === 0));
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_array_every_false() {
    let code = r#"
        let arr = [2, 3, 6];
        console.log(arr.every(x => x % 2 === 0));
    "#;
    assert_eq!(run_js_one(code), "false");
}

#[test]
fn test_array_find_index_found() {
    let code = r#"
        let arr = [10, 20, 30, 40];
        console.log(arr.findIndex(x => x >= 30));
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_array_find_index_not_found() {
    let code = r#"
        let arr = [10, 20, 30];
        console.log(arr.findIndex(x => x > 100));
    "#;
    assert_eq!(run_js_one(code), "-1");
}

// ============================================================
// Function.prototype.call / apply / hasOwnProperty
// ============================================================

#[test]
fn test_fn_call() {
    let code = r#"
        function greet(name) { return "hello " + name; }
        console.log(greet.call(null, "world"));
    "#;
    assert_eq!(run_js_one(code), "hello world");
}

#[test]
fn test_has_own_property() {
    let code = r#"
        let obj = { x: 1, y: 2 };
        console.log(obj.hasOwnProperty("x"));
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_has_own_property_missing() {
    let code = r#"
        let obj = { x: 1 };
        console.log(obj.hasOwnProperty("z"));
    "#;
    assert_eq!(run_js_one(code), "false");
}

// ============================================================
// Object static methods
// ============================================================

#[test]
fn test_object_keys() {
    let code = r#"
        let obj = { a: 1, b: 2 };
        let keys = Object.keys(obj);
        console.log(keys.length);
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_object_values() {
    let code = r#"
        let obj = { a: 10, b: 20 };
        let vals = Object.values(obj);
        console.log(vals.length);
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_object_entries() {
    let code = r#"
        let obj = { x: 1 };
        let entries = Object.entries(obj);
        console.log(entries.length);
    "#;
    assert_eq!(run_js_one(code), "1");
}

// ============================================================
// Array.from / Array.isArray
// ============================================================

#[test]
fn test_array_from() {
    let code = r#"
        let orig = [1, 2, 3];
        let copy = Array.from(orig);
        copy.push(4);
        console.log(orig.length);
    "#;
    // orig should still be 3 — copy is independent
    assert_eq!(run_js_one(code), "3");
}

#[test]
fn test_array_is_array() {
    assert_eq!(run_js_one("console.log(Array.isArray([1,2]))"), "true");
    assert_eq!(run_js_one("console.log(Array.isArray(42))"), "false");
}

// ============================================================
// VM opcode fixes: str_length with unicode
// ============================================================

#[test]
fn test_str_length_ascii() {
    assert_eq!(run_js_one(r#"console.log("hello".length)"#), "5");
}

#[test]
fn test_str_length_unicode() {
    assert_eq!(run_js_one(r#"console.log("héllo".length)"#), "5");
}

// ============================================================
// VM opcode fixes: f64_promote_f32
// ============================================================

#[test]
fn test_f64_promote_f32_not_noop() {
    // This is hard to test directly from JS, but verifying basic float ops still work
    assert_eq!(run_js_one("console.log(3.14 + 0)"), "3.14");
}
