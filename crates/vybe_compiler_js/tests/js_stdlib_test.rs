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
// MATH
// ============================================================

#[test]
fn test_math_floor() {
    assert_eq!(run_js_one("console.log(Math.floor(3.7))"), "3");
    assert_eq!(run_js_one("console.log(Math.floor(-1.2))"), "-2");
}

#[test]
fn test_math_ceil() {
    assert_eq!(run_js_one("console.log(Math.ceil(3.2))"), "4");
    assert_eq!(run_js_one("console.log(Math.ceil(-1.7))"), "-1");
}

#[test]
fn test_math_round() {
    assert_eq!(run_js_one("console.log(Math.round(3.5))"), "4");
    assert_eq!(run_js_one("console.log(Math.round(3.4))"), "3");
}

#[test]
fn test_math_abs() {
    assert_eq!(run_js_one("console.log(Math.abs(-5))"), "5");
    assert_eq!(run_js_one("console.log(Math.abs(5))"), "5");
}

#[test]
fn test_math_sqrt() {
    assert_eq!(run_js_one("console.log(Math.sqrt(9))"), "3");
    assert_eq!(run_js_one("console.log(Math.sqrt(2))"), "1.4142135623730951");
}

#[test]
fn test_math_pow() {
    assert_eq!(run_js_one("console.log(Math.pow(2, 10))"), "1024");
    assert_eq!(run_js_one("console.log(Math.pow(3, 3))"), "27");
}

#[test]
fn test_math_min_max() {
    assert_eq!(run_js_one("console.log(Math.min(3, 7))"), "3");
    assert_eq!(run_js_one("console.log(Math.max(3, 7))"), "7");
}

// ============================================================
// STRING METHODS
// ============================================================

#[test]
fn test_str_to_upper_lower() {
    assert_eq!(run_js_one(r#"console.log("hello".toUpperCase())"#), "HELLO");
    assert_eq!(run_js_one(r#"console.log("HELLO".toLowerCase())"#), "hello");
}

#[test]
fn test_str_trim() {
    assert_eq!(run_js_one(r#"console.log("  hello  ".trim())"#), "hello");
}

#[test]
fn test_str_slice() {
    assert_eq!(run_js_one(r#"console.log("hello world".slice(0, 5))"#), "hello");
    assert_eq!(run_js_one(r#"console.log("hello world".slice(6))"#), "world");
    assert_eq!(run_js_one(r#"console.log("hello".slice(-3))"#), "llo");
}

#[test]
fn test_str_index_of() {
    assert_eq!(run_js_one(r#"console.log("hello world".indexOf("world"))"#), "6");
    assert_eq!(run_js_one(r#"console.log("hello".indexOf("xyz"))"#), "-1");
}

#[test]
fn test_str_includes() {
    assert_eq!(run_js_one(r#"console.log("hello world".includes("world"))"#), "true");
    assert_eq!(run_js_one(r#"console.log("hello".includes("xyz"))"#), "false");
}

#[test]
fn test_str_split() {
    assert_eq!(run_js_one(r#"let parts = "a,b,c".split(","); console.log(parts[0], parts[1], parts[2])"#), "a b c");
}

#[test]
fn test_str_replace() {
    assert_eq!(run_js_one(r#"console.log("hello world".replace("world", "JS"))"#), "hello JS");
}

#[test]
fn test_str_starts_ends_with() {
    assert_eq!(run_js_one(r#"console.log("hello".startsWith("hel"))"#), "true");
    assert_eq!(run_js_one(r#"console.log("hello".endsWith("llo"))"#), "true");
    assert_eq!(run_js_one(r#"console.log("hello".startsWith("xyz"))"#), "false");
}

#[test]
fn test_str_char_at() {
    assert_eq!(run_js_one(r#"console.log("hello".charAt(1))"#), "e");
}

#[test]
fn test_str_substring() {
    assert_eq!(run_js_one(r#"console.log("hello world".substring(0, 5))"#), "hello");
}

// ============================================================
// ARRAY METHODS
// ============================================================

#[test]
fn test_arr_push() {
    assert_eq!(run_js_one("let a = [1, 2]; a.push(3); console.log(a)"), "1,2,3");
}

#[test]
fn test_arr_pop() {
    assert_eq!(run_js_one("let a = [1, 2, 3]; let x = a.pop(); console.log(x, a)"), "3 1,2");
}

#[test]
fn test_arr_shift() {
    assert_eq!(run_js_one("let a = [1, 2, 3]; let x = a.shift(); console.log(x, a)"), "1 2,3");
}

#[test]
fn test_arr_join() {
    assert_eq!(run_js_one(r#"console.log([1, 2, 3].join("-"))"#), "1-2-3");
    assert_eq!(run_js_one(r#"console.log([1, 2, 3].join())"#), "1,2,3");
}

#[test]
fn test_arr_reverse() {
    assert_eq!(run_js_one("let a = [1, 2, 3]; a.reverse(); console.log(a)"), "3,2,1");
}

#[test]
fn test_arr_concat() {
    assert_eq!(run_js_one("let a = [1, 2]; let b = [3, 4]; console.log(a.concat(b))"), "1,2,3,4");
}

#[test]
fn test_arr_slice() {
    assert_eq!(run_js_one("let a = [1, 2, 3, 4, 5]; console.log(a.slice(1, 3))"), "2,3");
}

// ============================================================
// GLOBAL FUNCTIONS
// ============================================================

#[test]
fn test_parse_int() {
    assert_eq!(run_js_one(r#"console.log(parseInt("42"))"#), "42");
    assert_eq!(run_js_one(r#"console.log(parseInt("3.14"))"#), "3");
}

#[test]
fn test_parse_float() {
    assert_eq!(run_js_one(r#"console.log(parseFloat("3.14"))"#), "3.14");
}

#[test]
fn test_is_nan() {
    assert_eq!(run_js_one(r#"console.log(isNaN(parseInt("abc")))"#), "true");
    assert_eq!(run_js_one("console.log(isNaN(42))"), "false");
}

// ============================================================
// JSON
// ============================================================

#[test]
fn test_json_stringify() {
    assert_eq!(run_js_one("console.log(JSON.stringify(42))"), "42");
    assert_eq!(run_js_one(r#"console.log(JSON.stringify("hello"))"#), r#""hello""#);
    assert_eq!(run_js_one("console.log(JSON.stringify([1, 2, 3]))"), "[1,2,3]");
    assert_eq!(run_js_one("console.log(JSON.stringify(null))"), "null");
}

// ============================================================
// COMBINED USAGE
// ============================================================

#[test]
fn test_combined_math_and_array() {
    let code = r#"
        let nums = [3.7, 1.2, 5.9, 2.1];
        let sum = 0;
        for (let i = 0; i < nums.length; i++) {
            sum = sum + Math.floor(nums[i]);
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "11"); // 3+1+5+2
}

#[test]
fn test_combined_string_split_loop() {
    let code = r#"
        let csv = "Alice,30,Bob,25,Charlie,35";
        let parts = csv.split(",");
        let names = [];
        for (let i = 0; i < parts.length; i += 2) {
            names.push(parts[i]);
        }
        console.log(names.join(" "));
    "#;
    assert_eq!(run_js_one(code), "Alice Bob Charlie");
}
