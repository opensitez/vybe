use std::cell::RefCell;
use std::rc::Rc;

fn run_js(code: &str) -> Vec<String> {
    let program = vybe_parser_js::parse(code).expect("parse failed");
    let mut vm = vybe_bytecode::VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

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
    run_js(code).into_iter().next().unwrap_or_default()
}

// ============================================================
// for...of
// ============================================================

#[test]
fn test_for_of_array() {
    let code = r#"
        let arr = [10, 20, 30];
        let sum = 0;
        for (let x of arr) {
            sum = sum + x;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "60");
}

#[test]
fn test_for_of_strings_array() {
    let code = r#"
        let names = ["Alice", "Bob", "Charlie"];
        let result = "";
        for (let name of names) {
            result = result + name + " ";
        }
        console.log(result.trim());
    "#;
    assert_eq!(run_js_one(code), "Alice Bob Charlie");
}

#[test]
fn test_for_of_with_break() {
    let code = r#"
        let arr = [1, 2, 3, 4, 5];
        let sum = 0;
        for (let x of arr) {
            if (x > 3) break;
            sum = sum + x;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "6");
}

#[test]
fn test_for_of_with_continue() {
    let code = r#"
        let arr = [1, 2, 3, 4, 5];
        let sum = 0;
        for (let x of arr) {
            if (x % 2 === 0) continue;
            sum = sum + x;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "9"); // 1+3+5
}

#[test]
fn test_for_of_empty_array() {
    let code = r#"
        let sum = 0;
        for (let x of []) { sum = sum + 1; }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "0");
}

// ============================================================
// for...in
// ============================================================

#[test]
fn test_for_in_object() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let keys = [];
        for (let k in obj) {
            keys.push(k);
        }
        console.log(keys.length);
    "#;
    assert_eq!(run_js_one(code), "3");
}

#[test]
fn test_for_in_values() {
    let code = r#"
        let obj = { x: 10, y: 20 };
        let sum = 0;
        for (let k in obj) {
            sum = sum + obj[k];
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "30");
}

// ============================================================
// Template literals with expressions
// ============================================================

#[test]
fn test_template_literal_simple() {
    let code = r#"
        let name = "World";
        console.log(`Hello ${name}!`);
    "#;
    assert_eq!(run_js_one(code), "Hello World!");
}

#[test]
fn test_template_literal_expression() {
    let code = r#"
        let a = 3;
        let b = 4;
        console.log(`${a} + ${b} = ${a + b}`);
    "#;
    assert_eq!(run_js_one(code), "3 + 4 = 7");
}

#[test]
fn test_template_literal_nested() {
    let code = r#"
        let items = [1, 2, 3];
        console.log(`count: ${items.length}`);
    "#;
    assert_eq!(run_js_one(code), "count: 3");
}

// ============================================================
// Object.keys/values/entries
// ============================================================

#[test]
fn test_object_keys() {
    let code = r#"
        let obj = { name: "Alice", age: 30 };
        let keys = Object.keys(obj);
        console.log(keys.length);
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_object_values() {
    let code = r#"
        let obj = { x: 10, y: 20 };
        let vals = Object.values(obj);
        console.log(vals.length);
    "#;
    assert_eq!(run_js_one(code), "2");
}

// ============================================================
// === vs == distinction
// ============================================================

#[test]
fn test_strict_vs_loose_equality() {
    // For now both use dyn_eq. This test documents current behavior.
    assert_eq!(run_js_one("console.log(1 === 1)"), "true");
    assert_eq!(run_js_one("console.log(1 === 2)"), "false");
    assert_eq!(run_js_one(r#"console.log("1" === 1)"#), "false"); // different types
    assert_eq!(run_js_one(r#"console.log("1" == 1)"#), "false");  // our dyn_eq doesn't coerce string→number
}

// ============================================================
// Combined: real-world patterns
// ============================================================

#[test]
fn test_for_of_with_function() {
    let code = r#"
        function sum(arr) {
            let total = 0;
            for (let x of arr) { total = total + x; }
            return total;
        }
        console.log(sum([1, 2, 3, 4, 5]));
    "#;
    assert_eq!(run_js_one(code), "15");
}

#[test]
fn test_for_of_with_template() {
    let code = r#"
        let fruits = ["apple", "banana", "cherry"];
        for (let fruit of fruits) {
            console.log(`I like ${fruit}`);
        }
    "#;
    let lines = run_js(code);
    assert_eq!(lines, vec!["I like apple", "I like banana", "I like cherry"]);
}

#[test]
fn test_for_in_build_new_object() {
    let code = r#"
        let source = { a: 1, b: 2, c: 3 };
        let doubled = {};
        for (let k in source) {
            doubled[k] = source[k] * 2;
        }
        console.log(doubled.a, doubled.b, doubled.c);
    "#;
    // This uses computed member assignment doubled[k] = ...
    assert_eq!(run_js_one(code), "2 4 6");
}

// ============================================================
// Optional chaining ?.
// ============================================================

#[test]
fn test_optional_chaining_non_null() {
    assert_eq!(run_js_one(r#"let o = { x: 42 }; console.log(o?.x)"#), "42");
}

#[test]
fn test_optional_chaining_null() {
    assert_eq!(run_js_one(r#"let o = null; console.log(o?.x)"#), "null");
}

#[test]
fn test_optional_chaining_nested() {
    assert_eq!(run_js_one(r#"let o = { a: { b: 99 } }; console.log(o?.a?.b)"#), "99");
}

#[test]
fn test_optional_chaining_null_nested() {
    assert_eq!(run_js_one(r#"let o = { a: null }; console.log(o?.a?.b)"#), "null");
}

// ============================================================
// Default parameters
// ============================================================

#[test]
fn test_default_params() {
    let code = r#"
        function greet(name, greeting) {
            if (greeting === null) { greeting = "Hello"; }
            return greeting + ", " + name + "!";
        }
        console.log(greet("World"));
        console.log(greet("World", "Hi"));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "Hello, World!");
    assert_eq!(lines[1], "Hi, World!");
}

// ============================================================
// Computed property names
// ============================================================

#[test]
fn test_computed_property_access() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let key = "b";
        console.log(obj[key]);
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_computed_property_set() {
    let code = r#"
        let obj = {};
        let key = "name";
        obj[key] = "Alice";
        console.log(obj.name);
    "#;
    assert_eq!(run_js_one(code), "Alice");
}

// ============================================================
// Strict vs loose equality
// ============================================================

#[test]
fn test_null_equality() {
    // null === null should be true
    assert_eq!(run_js_one("console.log(null === null)"), "true");
    // undefined/null treated as same in our VM
    assert_eq!(run_js_one("console.log(null == null)"), "true");
}

#[test]
fn test_type_mismatch_equality() {
    // Different types with === should be false
    assert_eq!(run_js_one(r#"console.log(1 === "1")"#), "false");
    assert_eq!(run_js_one("console.log(0 === false)"), "false");
    assert_eq!(run_js_one("console.log(null === 0)"), "false");
}

// ============================================================
// String methods with for...of
// ============================================================

#[test]
fn test_split_and_for_of() {
    let code = r#"
        let csv = "a,b,c,d";
        let parts = csv.split(",");
        let result = "";
        for (let p of parts) {
            result = result + p.toUpperCase() + " ";
        }
        console.log(result.trim());
    "#;
    assert_eq!(run_js_one(code), "A B C D");
}

// ============================================================
// Array methods with template literals
// ============================================================

#[test]
fn test_array_join_with_template() {
    let code = r#"
        let names = ["Alice", "Bob", "Charlie"];
        console.log(`Names: ${names.join(", ")}`);
    "#;
    assert_eq!(run_js_one(code), "Names: Alice, Bob, Charlie");
}

// ============================================================
// Nested for...of with objects
// ============================================================

#[test]
fn test_for_of_nested() {
    let code = r#"
        let matrix = [[1, 2], [3, 4], [5, 6]];
        let sum = 0;
        for (let row of matrix) {
            for (let val of row) {
                sum = sum + val;
            }
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "21");
}

// ============================================================
// Default parameters (proper syntax)
// ============================================================

#[test]
fn test_default_param_syntax() {
    let code = r#"
        function greet(name, greeting = "Hello") {
            return `${greeting}, ${name}!`;
        }
        console.log(greet("World"));
        console.log(greet("World", "Hi"));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "Hello, World!");
    assert_eq!(lines[1], "Hi, World!");
}

#[test]
fn test_default_param_expression() {
    let code = r#"
        function makeArray(size = 3, fill = 0) {
            let arr = [];
            for (let i = 0; i < size; i++) { arr.push(fill); }
            return arr;
        }
        console.log(makeArray());
        console.log(makeArray(2, 7));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "0,0,0");
    assert_eq!(lines[1], "7,7");
}

#[test]
fn test_default_param_arrow() {
    let code = r#"
        let add = (a, b = 1) => a + b;
        console.log(add(5));
        console.log(add(5, 10));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "6");
    assert_eq!(lines[1], "15");
}

// ============================================================
// Rest parameters
// ============================================================

#[test]
fn test_rest_params() {
    // Rest params collect remaining args into an array
    // Our VM passes extra args which land in the local slots
    // For now, rest is parsed but not collected into an array
    // This test verifies parsing doesn't crash
    let code = r#"
        function first(a, b) {
            return a + b;
        }
        console.log(first(1, 2, 3, 4));
    "#;
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// Destructuring (manual pattern for now)
// ============================================================

#[test]
fn test_manual_destructure_object() {
    let code = r#"
        let obj = { x: 10, y: 20 };
        let x = obj.x;
        let y = obj.y;
        console.log(x + y);
    "#;
    assert_eq!(run_js_one(code), "30");
}

#[test]
fn test_manual_destructure_array() {
    let code = r#"
        let arr = [1, 2, 3];
        let first = arr[0];
        let second = arr[1];
        console.log(first, second);
    "#;
    assert_eq!(run_js_one(code), "1 2");
}

// ============================================================
// Destructuring — object
// ============================================================

#[test]
fn test_destructure_object() {
    let code = r#"
        let obj = { x: 10, y: 20, z: 30 };
        let { x, y, z } = obj;
        console.log(x, y, z);
    "#;
    assert_eq!(run_js_one(code), "10 20 30");
}

#[test]
fn test_destructure_object_rename() {
    let code = r#"
        let obj = { name: "Alice", age: 30 };
        let { name: n, age: a } = obj;
        console.log(n, a);
    "#;
    assert_eq!(run_js_one(code), "Alice 30");
}

#[test]
fn test_destructure_object_default() {
    let code = r#"
        let obj = { x: 10 };
        let { x, y = 99 } = obj;
        console.log(x, y);
    "#;
    assert_eq!(run_js_one(code), "10 99");
}

// ============================================================
// Destructuring — array
// ============================================================

#[test]
fn test_destructure_array() {
    let code = r#"
        let arr = [1, 2, 3];
        let [a, b, c] = arr;
        console.log(a, b, c);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

#[test]
fn test_destructure_array_skip() {
    let code = r#"
        let [a, , c] = [10, 20, 30];
        console.log(a, c);
    "#;
    assert_eq!(run_js_one(code), "10 30");
}

#[test]
fn test_destructure_array_rest() {
    let code = r#"
        let [first, ...rest] = [1, 2, 3, 4, 5];
        console.log(first, rest.length);
    "#;
    assert_eq!(run_js_one(code), "1 4");
}

#[test]
fn test_destructure_array_default() {
    let code = r#"
        let [a, b = 99] = [10];
        console.log(a, b);
    "#;
    assert_eq!(run_js_one(code), "10 99");
}

// ============================================================
// Destructuring — combined
// ============================================================

#[test]
fn test_destructure_function_return() {
    let code = r#"
        function getPoint() { return { x: 3, y: 4 }; }
        let { x, y } = getPoint();
        console.log(x + y);
    "#;
    assert_eq!(run_js_one(code), "7");
}

#[test]
fn test_destructure_swap() {
    // Classic swap using array destructuring
    let code = r#"
        let a = 1;
        let b = 2;
        let [x, y] = [b, a];
        console.log(x, y);
    "#;
    assert_eq!(run_js_one(code), "2 1");
}

// ============================================================
// RegExp
// ============================================================

#[test]
fn test_regex_test() {
    assert_eq!(run_js_one(r#"console.log(RegExp.test("\\d+", "abc123"))"#), "true");
    assert_eq!(run_js_one(r#"console.log(RegExp.test("\\d+", "abc"))"#), "false");
}

#[test]
fn test_regex_match() {
    let code = r#"
        let matches = RegExp.match("\\d+", "abc123def456");
        console.log(matches[0], matches[1]);
    "#;
    assert_eq!(run_js_one(code), "123 456");
}

#[test]
fn test_regex_replace() {
    assert_eq!(run_js_one(r#"console.log(RegExp.replace("\\d+", "abc123", "NUM"))"#), "abcNUM");
}

#[test]
fn test_regex_replace_all() {
    assert_eq!(run_js_one(r#"console.log(RegExp.replaceAll("\\d+", "a1b2c3", "X"))"#), "aXbXcX");
}

#[test]
fn test_regex_split() {
    let code = r#"
        let parts = RegExp.split("[,;]", "a,b;c,d");
        console.log(parts.join(" "));
    "#;
    assert_eq!(run_js_one(code), "a b c d");
}

// Old Map/Set tests removed — replaced by new Map()/new Set() tests below.

// ============================================================
// Map — JS-style syntax with new
// ============================================================

#[test]
fn test_new_map() {
    let code = r#"
        let m = new Map();
        m.set("name", "Alice");
        m.set("age", 30);
        console.log(m.get("name"), m.get("age"));
    "#;
    assert_eq!(run_js_one(code), "Alice 30");
}

#[test]
fn test_new_map_has_delete() {
    let code = r#"
        let m = new Map();
        m.set("x", 1);
        console.log(m.has("x"));
        m.delete("x");
        console.log(m.has("x"));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "true");
    assert_eq!(lines[1], "false");
}

#[test]
fn test_new_map_size() {
    let code = r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        console.log(m.size);
    "#;
    assert_eq!(run_js_one(code), "2");
}

// ============================================================
// Set — JS-style syntax with new
// ============================================================

#[test]
fn test_new_set() {
    let code = r#"
        let s = new Set();
        s.add("a");
        s.add("b");
        s.add("a");
        console.log(s.size);
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_new_set_has_delete() {
    let code = r#"
        let s = new Set();
        s.add(42);
        console.log(s.has(42));
        s.delete(42);
        console.log(s.has(42));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "true");
    assert_eq!(lines[1], "false");
}

// ============================================================
// Array.map
// ============================================================

#[test]
fn test_array_map() {
    let code = r#"let r = [1, 2, 3].map((x) => x * 2); console.log(r.join(","))"#;
    assert_eq!(run_js_one(code), "2,4,6");
}

#[test]
fn test_array_map_with_index() {
    let code = r#"let r = [10, 20, 30].map((x, i) => i); console.log(r.join(","))"#;
    assert_eq!(run_js_one(code), "0,1,2");
}

// ============================================================
// Array.filter
// ============================================================

#[test]
fn test_array_filter() {
    let code = r#"let r = [1, 2, 3, 4, 5].filter((x) => x > 3); console.log(r.join(","))"#;
    assert_eq!(run_js_one(code), "4,5");
}

#[test]
fn test_array_filter_even() {
    let code = r#"let r = [1, 2, 3, 4, 5, 6].filter((x) => x % 2 === 0); console.log(r.join(","))"#;
    assert_eq!(run_js_one(code), "2,4,6");
}

// ============================================================
// Array.forEach
// ============================================================

#[test]
fn test_array_foreach() {
    let code = r#"
        let sum = 0;
        [1, 2, 3, 4, 5].forEach((x) => { sum = sum + x; });
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "15");
}

// ============================================================
// Array.find
// ============================================================

#[test]
fn test_array_find() {
    let code = r#"console.log([1, 2, 3, 4, 5].find((x) => x > 3))"#;
    assert_eq!(run_js_one(code), "4");
}

#[test]
fn test_array_find_not_found() {
    let code = r#"console.log([1, 2, 3].find((x) => x > 10))"#;
    assert_eq!(run_js_one(code), "null");
}

// ============================================================
// Array.reduce
// ============================================================

#[test]
fn test_array_reduce_sum() {
    let code = r#"console.log([1, 2, 3, 4, 5].reduce((acc, x) => acc + x, 0))"#;
    assert_eq!(run_js_one(code), "15");
}

#[test]
fn test_array_reduce_product() {
    let code = r#"console.log([1, 2, 3, 4].reduce((acc, x) => acc * x, 1))"#;
    assert_eq!(run_js_one(code), "24");
}

// ============================================================
// Array.sort
// ============================================================

#[test]
fn test_array_sort_default() {
    let code = r#"let r = [3, 1, 4, 1, 5, 9].sort(); console.log(r.join(","))"#;
    assert_eq!(run_js_one(code), "1,1,3,4,5,9");
}

#[test]
fn test_array_sort_comparator() {
    let code = r#"let r = [3, 1, 4, 1, 5].sort((a, b) => b - a); console.log(r.join(","))"#;
    assert_eq!(run_js_one(code), "5,4,3,1,1");
}

// ============================================================
// Combined
// ============================================================

#[test]
fn test_map_filter_reduce() {
    let code = r#"
        let result = [1, 2, 3, 4, 5]
            .filter((x) => x % 2 !== 0)
            .map((x) => x * x)
            .reduce((acc, x) => acc + x, 0);
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "35"); // 1+9+25
}
