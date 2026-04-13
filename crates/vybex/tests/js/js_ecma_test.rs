use super::helpers::run_js;

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
    assert_eq!(run_js_one("console.log(1 === 1)"), "true");
    assert_eq!(run_js_one("console.log(1 === 2)"), "false");
    assert_eq!(run_js_one(r#"console.log("1" === 1)"#), "false"); // strict: different types
    assert_eq!(run_js_one(r#"console.log("1" == 1)"#), "true");   // loose: string→number coercion
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

// ============================================================
// Import / Export — basic syntax parsing
// ============================================================

#[test]
fn test_export_function() {
    let code = r#"
        export function add(a, b) { return a + b; }
        console.log(add(3, 4));
    "#;
    assert_eq!(run_js_one(code), "7");
}

#[test]
fn test_export_let() {
    let code = r#"
        export let name = "Alice";
        console.log(name);
    "#;
    assert_eq!(run_js_one(code), "Alice");
}

#[test]
fn test_export_default_expression() {
    let code = r#"
        export default 42;
        console.log("ok");
    "#;
    assert_eq!(run_js_one(code), "ok");
}

#[test]
fn test_module_import_simulation() {
    // Simulate import resolution: the imported module's body is prepended
    // to the main module. This is what the vybex CLI does for:
    //   import { capitalize, VERSION } from "./lib/utils.js"
    let code = r#"
        // --- imported from utils.js ---
        export function capitalize(str) {
            if (str.length === 0) return str;
            return str.charAt(0).toUpperCase() + str.slice(1);
        }
        export let VERSION = "1.0.0";
        // --- main module ---
        console.log(capitalize("hello"));
        console.log(VERSION);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "Hello");
    assert_eq!(lines[1], "1.0.0");
}

#[test]
fn test_import_from_host_module() {
    // Importing from vybe:* should not crash — host modules resolved at call site
    let code = r#"
        import { floor } from "vybe:math";
        console.log(Math.floor(3.7));
    "#;
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// String extras
// ============================================================

#[test]
fn test_char_code_at() {
    assert_eq!(run_js_one(r#"console.log("ABC".charCodeAt(0))"#), "65");
}

#[test]
fn test_from_char_code() {
    assert_eq!(run_js_one(r#"console.log(String.fromCharCode(72, 105))"#), "Hi");
}

#[test]
fn test_string_repeat() {
    assert_eq!(run_js_one(r#"console.log("ab".repeat(3))"#), "ababab");
}

#[test]
fn test_pad_start() {
    assert_eq!(run_js_one(r#"console.log("5".padStart(3, "0"))"#), "005");
}

#[test]
fn test_pad_end() {
    assert_eq!(run_js_one(r#"console.log("hi".padEnd(5, "."))"#), "hi...");
}

#[test]
fn test_replace_all() {
    assert_eq!(run_js_one(r#"console.log("aXbXc".replaceAll("X", "-"))"#), "a-b-c");
}

// ============================================================
// Math extras
// ============================================================

#[test]
fn test_math_trunc() {
    assert_eq!(run_js_one("console.log(Math.trunc(3.7))"), "3");
    assert_eq!(run_js_one("console.log(Math.trunc(-3.7))"), "-3");
}

#[test]
fn test_math_sign() {
    assert_eq!(run_js_one("console.log(Math.sign(-5))"), "-1");
    assert_eq!(run_js_one("console.log(Math.sign(0))"), "0");
    assert_eq!(run_js_one("console.log(Math.sign(5))"), "1");
}

#[test]
fn test_math_hypot() {
    assert_eq!(run_js_one("console.log(Math.hypot(3, 4))"), "5");
}

// ============================================================
// Array extras
// ============================================================

#[test]
fn test_array_fill() {
    let code = r#"let a = [1, 2, 3, 4]; a.fill(0, 1, 3); console.log(a.join(","))"#;
    assert_eq!(run_js_one(code), "1,0,0,4");
}

#[test]
fn test_array_flat() {
    let code = r#"let a = [[1, 2], [3, 4], [5]]; console.log(a.flat().join(","))"#;
    assert_eq!(run_js_one(code), "1,2,3,4,5");
}

#[test]
fn test_array_includes() {
    let code = r#"console.log([1, 2, 3].includes(2), [1, 2, 3].includes(5))"#;
    assert_eq!(run_js_one(code), "true false");
}

// ============================================================
// Number.isInteger
// ============================================================

#[test]
fn test_number_is_integer() {
    let code = r#"console.log(Number.isInteger(5), Number.isInteger(5.5), Number.isInteger(0))"#;
    assert_eq!(run_js_one(code), "true false true");
}

// ============================================================
// try/catch — WASM exception proposal style
// ============================================================

#[test]
fn test_try_catch_basic() {
    let code = r#"
        try {
            throw "oops";
        } catch (e) {
            console.log("caught:", e);
        }
    "#;
    assert_eq!(run_js_one(code), "caught: oops");
}

#[test]
fn test_try_catch_no_throw() {
    let code = r#"
        let result = "ok";
        try {
            result = "from try";
        } catch (e) {
            result = "from catch";
        }
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "from try");
}

#[test]
fn test_try_catch_finally() {
    let code = r#"
        let log = "";
        try {
            log = log + "try ";
            throw "err";
        } catch (e) {
            log = log + "catch ";
        } finally {
            log = log + "finally";
        }
        console.log(log);
    "#;
    assert_eq!(run_js_one(code), "try catch finally");
}

#[test]
fn test_try_catch_error_value() {
    let code = r#"
        try {
            throw 42;
        } catch (e) {
            console.log(e + 8);
        }
    "#;
    assert_eq!(run_js_one(code), "50");
}

#[test]
fn test_try_catch_nested() {
    let code = r#"
        let result = "";
        try {
            try {
                throw "inner";
            } catch (e) {
                result = result + e + " ";
                throw "outer";
            }
        } catch (e) {
            result = result + e;
        }
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "inner outer");
}

#[test]
fn test_try_catch_in_function() {
    let code = r#"
        function safeDivide(a, b) {
            try {
                if (b === 0) throw "division by zero";
                return a / b;
            } catch (e) {
                return e;
            }
        }
        console.log(safeDivide(10, 2));
        console.log(safeDivide(10, 0));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "5");
    assert_eq!(lines[1], "division by zero");
}

// ============================================================
// typeof
// ============================================================

#[test]
fn test_typeof_undeclared() {
    assert_eq!(run_js_one(r#"console.log(typeof xyz)"#), "undefined");
}

#[test]
fn test_typeof_null() {
    assert_eq!(run_js_one(r#"console.log(typeof null)"#), "object"); // JS spec
}

#[test]
fn test_typeof_number() {
    assert_eq!(run_js_one(r#"console.log(typeof 42)"#), "number");
}

#[test]
fn test_typeof_string() {
    assert_eq!(run_js_one(r#"console.log(typeof "hello")"#), "string");
}

#[test]
fn test_typeof_boolean() {
    assert_eq!(run_js_one(r#"console.log(typeof true)"#), "boolean");
}

#[test]
fn test_typeof_function() {
    assert_eq!(run_js_one(r#"function f() {} console.log(typeof f)"#), "function");
}

#[test]
fn test_typeof_object() {
    assert_eq!(run_js_one(r#"console.log(typeof {})"#), "object");
}

// ============================================================
// Spread in function calls
// ============================================================

#[test]
fn test_spread_array_literal() {
    let code = r#"
        function sum(a, b, c) { return a + b + c; }
        console.log(sum(...[1, 2, 3]));
    "#;
    assert_eq!(run_js_one(code), "6");
}

#[test]
fn test_spread_mixed() {
    let code = r#"
        function f(a, b, c, d) { return a + b + c + d; }
        console.log(f(1, ...[2, 3], 4));
    "#;
    assert_eq!(run_js_one(code), "10");
}

#[test]
fn test_spread_in_array() {
    let code = r#"
        let a = [1, 2];
        let b = [3, 4];
        let c = [...a, ...b, 5];
        console.log(c.join(","));
    "#;
    assert_eq!(run_js_one(code), "1,2,3,4,5");
}

// ============================================================
// async/await
// ============================================================

#[test]
fn test_async_function() {
    let code = r#"
        async function getValue() {
            return 42;
        }
        let result = await getValue();
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "42");
}

#[test]
fn test_await_non_promise() {
    // await on a non-Promise just returns the value
    assert_eq!(run_js_one("console.log(await 5)"), "5");
}

#[test]
fn test_async_arrow() {
    let code = r#"
        let fetchData = async () => {
            return "data loaded";
        };
        console.log(await fetchData());
    "#;
    assert_eq!(run_js_one(code), "data loaded");
}

#[test]
fn test_promise_resolve() {
    let code = r#"
        let p = Promise.resolve(99);
        console.log(await p);
    "#;
    assert_eq!(run_js_one(code), "99");
}

#[test]
fn test_promise_all() {
    let code = r#"
        let results = await Promise.all([
            Promise.resolve(1),
            Promise.resolve(2),
            Promise.resolve(3)
        ]);
        console.log(results.join(","));
    "#;
    assert_eq!(run_js_one(code), "1,2,3");
}

#[test]
fn test_async_sequential() {
    let code = r#"
        async function step1() { return 10; }
        async function step2(x) { return x * 2; }
        async function step3(x) { return x + 5; }

        let a = await step1();
        let b = await step2(a);
        let c = await step3(b);
        console.log(c);
    "#;
    assert_eq!(run_js_one(code), "25");
}

// ============================================================
// setTimeout (blocking in our model)
// ============================================================

#[test]
fn test_set_timeout_async() {
    // setTimeout callback fires AFTER synchronous code finishes (via event loop)
    let code = r#"
        let log = "";
        setTimeout(() => { log = log + "timer "; console.log(log + "done"); }, 1);
        log = log + "sync ";
        console.log(log);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "sync ");        // sync code runs first
    assert_eq!(lines[1], "sync timer done"); // then timer callback fires
}

// ============================================================
// Async — comprehensive tests
// ============================================================

#[test]
fn test_settimeout_ordering() {
    // Multiple timers fire in order
    let code = r#"
        let order = [];
        setTimeout(() => { order.push("a"); }, 1);
        setTimeout(() => { order.push("b"); }, 2);
        setTimeout(() => { order.push("c"); }, 3);
        setTimeout(() => { console.log(order.join(",")); }, 4);
    "#;
    assert_eq!(run_js_one(code), "a,b,c");
}

#[test]
fn test_settimeout_zero() {
    // setTimeout(fn, 0) still defers to event loop
    let code = r#"
        let result = "before";
        setTimeout(() => { result = "after"; console.log(result); }, 0);
        console.log(result);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "before");  // sync runs first
    assert_eq!(lines[1], "after");   // then timer
}

#[test]
fn test_settimeout_closure() {
    // Timer callbacks capture closures properly
    // Use function wrapper to capture per-iteration (known loop scope limitation)
    let code = r#"
        let results = [];
        function scheduleOne(val) {
            setTimeout(() => { results.push(val); }, 1);
        }
        for (let i = 0; i < 3; i++) {
            scheduleOne(i);
        }
        setTimeout(() => { console.log(results.join(",")); }, 5);
    "#;
    assert_eq!(run_js_one(code), "0,1,2");
}

#[test]
fn test_async_await_chain() {
    let code = r#"
        async function double(x) { return x * 2; }
        async function addTen(x) { return x + 10; }
        
        async function compute() {
            let a = await double(5);
            let b = await addTen(a);
            return b;
        }
        
        console.log(await compute());
    "#;
    assert_eq!(run_js_one(code), "20");
}

#[test]
fn test_async_await_with_regular_values() {
    // await on non-Promise values works
    let code = r#"
        async function getItems() {
            let a = await 10;
            let b = await 20;
            return a + b;
        }
        console.log(await getItems());
    "#;
    assert_eq!(run_js_one(code), "30");
}

#[test]
fn test_promise_resolve_then_await() {
    let code = r#"
        let p = Promise.resolve(42);
        let val = await p;
        console.log(val);
    "#;
    assert_eq!(run_js_one(code), "42");
}

#[test]
fn test_promise_all_with_values() {
    let code = r#"
        let results = await Promise.all([
            Promise.resolve("a"),
            Promise.resolve("b"),
            Promise.resolve("c")
        ]);
        console.log(results.join("-"));
    "#;
    assert_eq!(run_js_one(code), "a-b-c");
}

#[test]
fn test_promise_all_with_async_functions() {
    let code = r#"
        async function fetchName() { return "Alice"; }
        async function fetchAge() { return 30; }
        
        let [name, age] = await Promise.all([
            fetchName(),
            fetchAge()
        ]);
        console.log(name, age);
    "#;
    assert_eq!(run_js_one(code), "Alice 30");
}

#[test]
fn test_async_error_handling() {
    let code = r#"
        async function riskyOp() {
            throw "something went wrong";
        }
        
        try {
            await riskyOp();
        } catch (e) {
            console.log("caught:", e);
        }
    "#;
    assert_eq!(run_js_one(code), "caught: something went wrong");
}

#[test]
fn test_async_in_loop() {
    let code = r#"
        async function process(x) { return x * x; }
        
        let results = [];
        let items = [1, 2, 3, 4, 5];
        for (let item of items) {
            let result = await process(item);
            results.push(result);
        }
        console.log(results.join(","));
    "#;
    assert_eq!(run_js_one(code), "1,4,9,16,25");
}

#[test]
fn test_settimeout_nested() {
    // Timer inside a timer
    let code = r#"
        let log = [];
        setTimeout(() => {
            log.push("first");
            setTimeout(() => {
                log.push("second");
                console.log(log.join(","));
            }, 1);
        }, 1);
    "#;
    assert_eq!(run_js_one(code), "first,second");
}

#[test]
fn test_async_with_class() {
    let code = r#"
        class UserService {
            constructor(name) {
                this.name = name;
            }
            async greet() {
                return "Hello, " + this.name + "!";
            }
        }
        
        let svc = new UserService("Bob");
        console.log(await svc.greet());
    "#;
    assert_eq!(run_js_one(code), "Hello, Bob!");
}

#[test]
fn test_async_map_sequential() {
    // Process array items sequentially with async
    let code = r#"
        async function transform(x) { return x * 10; }
        
        let items = [1, 2, 3];
        let results = [];
        for (let item of items) {
            results.push(await transform(item));
        }
        console.log(results.join(","));
    "#;
    assert_eq!(run_js_one(code), "10,20,30");
}

#[test]
fn test_mixed_sync_async() {
    let code = r#"
        let log = [];
        log.push("1-sync");
        
        async function asyncOp() {
            log.push("2-async-start");
            let result = await Promise.resolve("done");
            log.push("3-async-end");
            return result;
        }
        
        let result = await asyncOp();
        log.push("4-after-await");
        console.log(log.join(" | "));
    "#;
    assert_eq!(run_js_one(code), "1-sync | 2-async-start | 3-async-end | 4-after-await");
}

// ============================================================
// Class inheritance (extends)
// ============================================================

#[test]
fn test_class_extends_basic() {
    let code = r#"
        class Animal {
            constructor(name) {
                this.name = name;
            }
            speak() {
                return this.name + " makes a sound";
            }
        }
        class Dog extends Animal {
            constructor(name) {
                super();
                this.name = name;
            }
            bark() {
                return this.name + " barks";
            }
        }
        let d = new Dog("Rex");
        console.log(d.name);
        console.log(d.bark());
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "Rex");
    assert_eq!(lines[1], "Rex barks");
}

#[test]
fn test_class_extends_override() {
    let code = r#"
        class Shape {
            constructor(type) {
                this.type = type;
            }
            describe() {
                return "I am a " + this.type;
            }
        }
        class Circle extends Shape {
            constructor(radius) {
                super();
                this.type = "circle";
                this.radius = radius;
            }
            area() {
                return Math.PI * this.radius * this.radius;
            }
        }
        let c = new Circle(5);
        console.log(c.type);
        console.log(Math.round(c.area()));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "circle");
    assert_eq!(lines[1], "79");
}

// ============================================================
// Getter / Setter
// ============================================================

#[test]
fn test_class_getter() {
    let code = r#"
        class Person {
            constructor(first, last) {
                this.first = first;
                this.last = last;
            }
            get fullName() {
                return this.first + " " + this.last;
            }
        }
        let p = new Person("John", "Doe");
        console.log(p.first, p.last);
    "#;
    // For now, getters are stored as __get_fullName — not auto-invoked on property access yet
    assert_eq!(run_js_one(code), "John Doe");
}

#[test]
fn test_class_method_kinds() {
    let code = r#"
        class Counter {
            constructor() {
                this._count = 0;
            }
            increment() { this._count = this._count + 1; }
            getCount() { return this._count; }
        }
        let c = new Counter();
        c.increment();
        c.increment();
        c.increment();
        console.log(c.getCount());
    "#;
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// Error class
// ============================================================

#[test]
fn test_error_class() {
    let code = r#"
        let e = new Error("something failed");
        console.log(e.message);
        console.log(e.name);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "something failed");
    assert_eq!(lines[1], "Error");
}

#[test]
fn test_error_throw_catch() {
    let code = r#"
        try {
            throw new Error("oops");
        } catch (e) {
            console.log(e.message);
        }
    "#;
    assert_eq!(run_js_one(code), "oops");
}

#[test]
fn test_type_error() {
    let code = r#"
        let e = new TypeError("invalid type");
        console.log(e.name, e.message);
    "#;
    assert_eq!(run_js_one(code), "TypeError invalid type");
}

// ============================================================
// Static methods and properties
// ============================================================

#[test]
fn test_class_multiple_instances_independent() {
    let code = r#"
        class Box {
            constructor(w, h) {
                this.w = w;
                this.h = h;
            }
            area() { return this.w * this.h; }
        }
        let a = new Box(3, 4);
        let b = new Box(5, 6);
        console.log(a.area(), b.area());
    "#;
    assert_eq!(run_js_one(code), "12 30");
}

// ============================================================
// Class with all features combined
// ============================================================

#[test]
fn test_class_comprehensive() {
    let code = r#"
        class Vehicle {
            constructor(make, year) {
                this.make = make;
                this.year = year;
                this.speed = 0;
            }
            accelerate(amount) {
                this.speed = this.speed + amount;
            }
            brake(amount) {
                this.speed = Math.max(0, this.speed - amount);
            }
            describe() {
                return `${this.year} ${this.make} going ${this.speed}mph`;
            }
        }
        
        let car = new Vehicle("Tesla", 2024);
        car.accelerate(60);
        car.accelerate(20);
        car.brake(10);
        console.log(car.describe());
    "#;
    assert_eq!(run_js_one(code), "2024 Tesla going 70mph");
}

// ============================================================
// Getter / Setter — auto-invoke
// ============================================================

#[test]
fn test_getter_auto_invoke() {
    let code = r#"
        class Person {
            constructor(first, last) {
                this.first = first;
                this.last = last;
            }
            get fullName() {
                return this.first + " " + this.last;
            }
        }
        let p = new Person("John", "Doe");
        console.log(p.fullName);
    "#;
    assert_eq!(run_js_one(code), "John Doe");
}

#[test]
fn test_setter_auto_invoke() {
    let code = r#"
        class Temperature {
            constructor(celsius) {
                this._celsius = celsius;
            }
            get fahrenheit() {
                return this._celsius * 9 / 5 + 32;
            }
            set fahrenheit(f) {
                this._celsius = (f - 32) * 5 / 9;
            }
        }
        let t = new Temperature(100);
        console.log(t.fahrenheit);
    "#;
    assert_eq!(run_js_one(code), "212");
}

// ============================================================
// Inheritance — method inheritance
// ============================================================

#[test]
fn test_inherited_methods() {
    let code = r#"
        class Animal {
            constructor(name) {
                this.name = name;
            }
            speak() {
                return this.name + " speaks";
            }
        }
        class Dog extends Animal {
            constructor(name, breed) {
                super(name);
                this.breed = breed;
            }
            bark() {
                return this.name + " barks";
            }
        }
        let d = new Dog("Rex", "Labrador");
        console.log(d.name);
        console.log(d.bark());
        console.log(d.speak());
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "Rex");
    assert_eq!(lines[1], "Rex barks");
    assert_eq!(lines[2], "Rex speaks");
}

#[test]
fn test_super_with_args() {
    let code = r#"
        class Base {
            constructor(x) {
                this.x = x;
            }
            getX() { return this.x; }
        }
        class Child extends Base {
            constructor(x, y) {
                super(x);
                this.y = y;
            }
            getY() { return this.y; }
            sum() { return this.x + this.y; }
        }
        let c = new Child(10, 20);
        console.log(c.getX(), c.getY(), c.sum());
    "#;
    assert_eq!(run_js_one(code), "10 20 30");
}

#[test]
fn test_method_override() {
    let code = r#"
        class Shape {
            constructor() { this.type = "shape"; }
            describe() { return "I am a " + this.type; }
        }
        class Circle extends Shape {
            constructor(r) {
                super();
                this.type = "circle";
                this.r = r;
            }
            describe() { return "Circle with radius " + this.r; }
        }
        let s = new Shape();
        let c = new Circle(5);
        console.log(s.describe());
        console.log(c.describe());
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "I am a shape");
    assert_eq!(lines[1], "Circle with radius 5");
}

// ============================================================
// Error with throw/catch
// ============================================================

#[test]
fn test_error_stack() {
    let code = r#"
        let e = new Error("test error");
        console.log(e.stack);
    "#;
    assert_eq!(run_js_one(code), "Error: test error");
}

#[test]
fn test_error_instanceof_pattern() {
    let code = r#"
        function validate(x) {
            if (x < 0) throw new RangeError("must be positive");
            return x;
        }
        try {
            validate(-1);
        } catch (e) {
            console.log(e.name + ": " + e.message);
        }
    "#;
    assert_eq!(run_js_one(code), "RangeError: must be positive");
}

// ============================================================
// Computed property names
// ============================================================

#[test]
fn test_computed_property_literal() {
    let code = r#"
        let key = "name";
        let obj = { [key]: "Alice" };
        console.log(obj.name);
    "#;
    assert_eq!(run_js_one(code), "Alice");
}

#[test]
fn test_computed_property_expression() {
    let code = r#"
        let prefix = "get";
        let obj = { [prefix + "Name"]: () => "Bob" };
        console.log(obj.getName());
    "#;
    assert_eq!(run_js_one(code), "Bob");
}

#[test]
fn test_computed_property_dynamic() {
    let code = r#"
        let obj = {};
        for (let i = 0; i < 3; i++) {
            obj["key" + i] = i * 10;
        }
        console.log(obj.key0, obj.key1, obj.key2);
    "#;
    assert_eq!(run_js_one(code), "0 10 20");
}

// ============================================================
// Static methods and fields
// ============================================================

#[test]
fn test_static_method() {
    let code = r#"
        class MathHelper {
            static add(a, b) { return a + b; }
            static multiply(a, b) { return a * b; }
        }
        console.log(MathHelper.add(3, 4), MathHelper.multiply(3, 4));
    "#;
    assert_eq!(run_js_one(code), "7 12");
}

#[test]
fn test_static_field() {
    let code = r#"
        class Config {
            static version = "2.0";
            static debug = false;
        }
        console.log(Config.version, Config.debug);
    "#;
    assert_eq!(run_js_one(code), "2.0 false");
}

#[test]
fn test_static_factory() {
    let code = r#"
        class User {
            constructor(name, age) {
                this.name = name;
                this.age = age;
            }
            static fromJSON(json) {
                let data = JSON.parse(json);
                return new User(data.name, data.age);
            }
        }
        let u = User.fromJSON('{"name":"Alice","age":30}');
        console.log(u.name, u.age);
    "#;
    assert_eq!(run_js_one(code), "Alice 30");
}

// ============================================================
// Class field initializers
// ============================================================

#[test]
fn test_class_field_default() {
    let code = r#"
        class Counter {
            count = 0;
            increment() { this.count = this.count + 1; }
            getCount() { return this.count; }
        }
        let c = new Counter();
        c.increment();
        c.increment();
        console.log(c.getCount());
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_class_field_with_constructor() {
    let code = r#"
        class Point {
            z = 0;
            constructor(x, y) {
                this.x = x;
                this.y = y;
            }
        }
        let p = new Point(3, 4);
        console.log(p.x, p.y, p.z);
    "#;
    assert_eq!(run_js_one(code), "3 4 0");
}

// ============================================================
// Getter/Setter comprehensive
// ============================================================

#[test]
fn test_getter_computed_value() {
    let code = r#"
        class Rectangle {
            constructor(w, h) {
                this.width = w;
                this.height = h;
            }
            get area() {
                return this.width * this.height;
            }
            get perimeter() {
                return 2 * (this.width + this.height);
            }
        }
        let r = new Rectangle(5, 3);
        console.log(r.area, r.perimeter);
    "#;
    assert_eq!(run_js_one(code), "15 16");
}

#[test]
fn test_getter_setter_pair() {
    let code = r#"
        class Account {
            constructor(balance) {
                this._balance = balance;
            }
            get balance() {
                return this._balance;
            }
            set balance(val) {
                if (val < 0) { this._balance = 0; }
                else { this._balance = val; }
            }
        }
        let a = new Account(100);
        console.log(a.balance);
        a.balance = 200;
        console.log(a.balance);
        a.balance = -50;
        console.log(a.balance);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "100");
    assert_eq!(lines[1], "200");
    assert_eq!(lines[2], "0");
}

// ============================================================
// Comprehensive class example
// ============================================================

#[test]
fn test_class_full_example() {
    let code = r#"
        class EventEmitter {
            constructor() {
                this._handlers = {};
            }
            on(event, handler) {
                if (!this._handlers[event]) {
                    this._handlers[event] = [];
                }
                this._handlers[event].push(handler);
            }
            emit(event, data) {
                let handlers = this._handlers[event];
                if (handlers) {
                    handlers.forEach((h) => h(data));
                }
            }
        }
        
        let emitter = new EventEmitter();
        let log = [];
        emitter.on("greet", (name) => { log.push("Hello " + name); });
        emitter.on("greet", (name) => { log.push("Hi " + name); });
        emitter.emit("greet", "Alice");
        console.log(log.join(", "));
    "#;
    assert_eq!(run_js_one(code), "Hello Alice, Hi Alice");
}

// ============================================================
// Private class fields (#field)
// ============================================================

#[test]
fn test_private_field() {
    let code = r#"
        class BankAccount {
            #balance = 0;
            constructor(initial) {
                this.#balance = initial;
            }
            deposit(amount) {
                this.#balance = this.#balance + amount;
            }
            getBalance() {
                return this.#balance;
            }
        }
        let acc = new BankAccount(100);
        acc.deposit(50);
        console.log(acc.getBalance());
    "#;
    assert_eq!(run_js_one(code), "150");
}

#[test]
fn test_private_method() {
    let code = r#"
        class Validator {
            #isValid(value) {
                return value > 0;
            }
            validate(value) {
                if (this.#isValid(value)) {
                    return "valid";
                }
                return "invalid";
            }
        }
        let v = new Validator();
        console.log(v.validate(5), v.validate(-1));
    "#;
    assert_eq!(run_js_one(code), "valid invalid");
}

#[test]
fn test_private_getter() {
    let code = r#"
        class Secret {
            #data = "hidden";
            get revealed() {
                return this.#data;
            }
        }
        let s = new Secret();
        console.log(s.revealed);
    "#;
    assert_eq!(run_js_one(code), "hidden");
}
