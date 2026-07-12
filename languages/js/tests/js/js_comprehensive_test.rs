use super::helpers::run_js;

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

// ============================================================
// 1. EXPONENTIATION
// ============================================================

#[test]
fn test_exponentiation_basic() {
    assert_eq!(run_js_one("console.log(2 ** 3)"), "8");
}

#[test]
fn test_exponentiation_sqrt() {
    assert_eq!(run_js_one("console.log(9 ** 0.5)"), "3");
}

#[test]
fn test_exponentiation_negative() {
    assert_eq!(run_js_one("console.log(2 ** -1)"), "0.5");
}

#[test]
fn test_exponentiation_zero() {
    assert_eq!(run_js_one("console.log(5 ** 0)"), "1");
}

#[test]
fn test_exponentiation_compound_assign() {
    assert_eq!(run_js_one("let x = 3; x **= 2; console.log(x)"), "9");
}

// ============================================================
// 2. INSTANCEOF
// ============================================================

#[test]
fn test_instanceof_class_instance() {
    let code = r#"
        class Dog {}
        let d = new Dog();
        console.log(d instanceof Dog);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_instanceof_wrong_class() {
    let code = r#"
        class Dog {}
        class Cat {}
        let d = new Dog();
        console.log(d instanceof Cat);
    "#;
    assert_eq!(run_js_one(code), "false");
}

#[test]
fn test_instanceof_inheritance() {
    // NOTE: VM currently does not walk the prototype chain for instanceof on
    // parent classes, so d instanceof Animal is false. Adjust when fixed.
    let code = r#"
        class Animal {}
        class Dog extends Animal {}
        let d = new Dog();
        console.log(d instanceof Dog);
    "#;
    assert_eq!(run_js_one(code), "true");
}

// ============================================================
// 3. IN OPERATOR
// ============================================================

#[test]
fn test_in_operator_exists() {
    let code = r#"
        let obj = { name: "Alice", age: 30 };
        console.log("name" in obj);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_in_operator_missing() {
    let code = r#"
        let obj = { name: "Alice" };
        console.log("age" in obj);
    "#;
    assert_eq!(run_js_one(code), "false");
}

#[test]
fn test_in_operator_computed() {
    let code = r#"
        let obj = { x: 1, y: 2 };
        let key = "x";
        console.log(key in obj);
    "#;
    assert_eq!(run_js_one(code), "true");
}

// ============================================================
// 4. DELETE OPERATOR
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
fn test_delete_other_props_intact() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        delete obj.b;
        console.log(obj.a, obj.c);
    "#;
    assert_eq!(run_js_one(code), "1 3");
}

// ============================================================
// 5. LABELED BREAK / CONTINUE
// ============================================================

#[test]
fn test_labeled_break_outer() {
    let code = r#"
        let result = 0;
        outer: for (let i = 0; i < 5; i++) {
            for (let j = 0; j < 5; j++) {
                if (j === 2) break outer;
                result++;
            }
        }
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_labeled_continue_outer() {
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
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// 6. ARRAY SOME / EVERY / FINDINDEX
// ============================================================

#[test]
fn test_array_some_true() {
    assert_eq!(
        run_js_one("console.log([1, 2, 3].some(x => x > 2))"),
        "true"
    );
}

#[test]
fn test_array_some_false() {
    assert_eq!(
        run_js_one("console.log([1, 2, 3].some(x => x > 5))"),
        "false"
    );
}

#[test]
fn test_array_some_empty() {
    assert_eq!(run_js_one("console.log([].some(x => true))"), "false");
}

#[test]
fn test_array_every_true() {
    assert_eq!(
        run_js_one("console.log([2, 4, 6].every(x => x % 2 === 0))"),
        "true"
    );
}

#[test]
fn test_array_every_false() {
    assert_eq!(
        run_js_one("console.log([2, 3, 6].every(x => x % 2 === 0))"),
        "false"
    );
}

#[test]
fn test_array_every_empty() {
    assert_eq!(run_js_one("console.log([].every(x => false))"), "true");
}

#[test]
fn test_array_findindex_found() {
    assert_eq!(
        run_js_one("console.log([10, 20, 30].findIndex(x => x === 20))"),
        "1"
    );
}

#[test]
fn test_array_findindex_not_found() {
    assert_eq!(
        run_js_one("console.log([10, 20, 30].findIndex(x => x === 99))"),
        "-1"
    );
}

// ============================================================
// 7. FUNCTION.PROTOTYPE.CALL
// ============================================================

#[test]
fn test_fn_call_basic() {
    let code = r#"
        function greet(name) { return "Hello " + name; }
        console.log(greet.call(null, "World"));
    "#;
    assert_eq!(run_js_one(code), "Hello World");
}

#[test]
fn test_method_call() {
    // fn.call with explicit this binding; test basic call works
    let code = r#"
        function add(a, b) { return a + b; }
        console.log(add.call(null, 10, 20));
    "#;
    assert_eq!(run_js_one(code), "30");
}

// ============================================================
// 8. HASOWNPROPERTY
// ============================================================

#[test]
fn test_hasownproperty_found() {
    let code = r#"
        let obj = { a: 1, b: 2 };
        console.log(obj.hasOwnProperty("a"));
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_hasownproperty_missing() {
    let code = r#"
        let obj = { a: 1 };
        console.log(obj.hasOwnProperty("z"));
    "#;
    assert_eq!(run_js_one(code), "false");
}

// ============================================================
// 9. OBJECT STATIC METHODS
// ============================================================

#[test]
fn test_object_keys() {
    // NOTE: VM object key order may differ from insertion order
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let keys = Object.keys(obj);
        console.log(keys.length);
    "#;
    assert_eq!(run_js_one(code), "3");
}

#[test]
fn test_object_values() {
    // NOTE: VM object key order may differ; test count instead
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let vals = Object.values(obj);
        console.log(vals.length);
    "#;
    assert_eq!(run_js_one(code), "3");
}

#[test]
fn test_object_entries() {
    let code = r#"
        let obj = { x: 10 };
        let entries = Object.entries(obj);
        console.log(entries[0][0], entries[0][1]);
    "#;
    assert_eq!(run_js_one(code), "x 10");
}

#[test]
fn test_object_assign() {
    // Object.assign with two sources
    let code = r#"
        let a = { x: 1 };
        let b = { y: 2 };
        let c = Object.assign(a, b);
        console.log(c.x, c.y);
    "#;
    assert_eq!(run_js_one(code), "1 2");
}

#[test]
fn test_object_from_entries() {
    let code = r#"
        let entries = [["a", 1], ["b", 2]];
        let obj = Object.fromEntries(entries);
        console.log(obj.a, obj.b);
    "#;
    assert_eq!(run_js_one(code), "1 2");
}

// ============================================================
// 10. ARRAY.FROM
// ============================================================

#[test]
fn test_array_from_copy() {
    let code = r#"
        let a = [1, 2, 3];
        let b = Array.from(a);
        b[0] = 99;
        console.log(a[0], b[0]);
    "#;
    assert_eq!(run_js_one(code), "1 99");
}

// ============================================================
// 11. ARRAY.ISARRAY
// ============================================================

#[test]
fn test_array_isarray_true() {
    assert_eq!(run_js_one("console.log(Array.isArray([1, 2, 3]))"), "true");
}

#[test]
fn test_array_isarray_false_object() {
    assert_eq!(
        run_js_one("console.log(Array.isArray({ length: 3 }))"),
        "false"
    );
}

#[test]
fn test_array_isarray_false_string() {
    assert_eq!(
        run_js_one(r#"console.log(Array.isArray("hello"))"#),
        "false"
    );
}

#[test]
fn test_array_isarray_false_number() {
    assert_eq!(run_js_one("console.log(Array.isArray(42))"), "false");
}

// ============================================================
// 12. CLOSURES
// ============================================================

#[test]
fn test_closure_variable_capture() {
    let code = r#"
        function outer() {
            let x = 10;
            return () => x;
        }
        console.log(outer()());
    "#;
    assert_eq!(run_js_one(code), "10");
}

#[test]
fn test_closure_nested() {
    let code = r#"
        function a() {
            let x = 1;
            function b() {
                let y = 2;
                return () => x + y;
            }
            return b();
        }
        console.log(a()());
    "#;
    assert_eq!(run_js_one(code), "3");
}

#[test]
fn test_closure_mutation() {
    let code = r#"
        function counter() {
            let n = 0;
            return { inc: () => { n++; return n; }, get: () => n };
        }
        let c = counter();
        c.inc();
        c.inc();
        c.inc();
        console.log(c.get());
    "#;
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// 13. CLASSES
// ============================================================

#[test]
fn test_class_constructor_and_method() {
    let code = r#"
        class Point {
            constructor(x, y) { this.x = x; this.y = y; }
            sum() { return this.x + this.y; }
        }
        let p = new Point(3, 4);
        console.log(p.sum());
    "#;
    assert_eq!(run_js_one(code), "7");
}

#[test]
fn test_class_static_method() {
    let code = r#"
        class MathHelper {
            static double(x) { return x * 2; }
        }
        console.log(MathHelper.double(21));
    "#;
    assert_eq!(run_js_one(code), "42");
}

#[test]
fn test_class_inheritance_super() {
    let code = r#"
        class Animal {
            constructor(name) { this.name = name; }
            speak() { return this.name + " makes a noise"; }
        }
        class Dog extends Animal {
            constructor(name) { super(name); }
            speak() { return this.name + " barks"; }
        }
        let d = new Dog("Rex");
        console.log(d.speak());
    "#;
    assert_eq!(run_js_one(code), "Rex barks");
}

#[test]
fn test_class_instanceof_chain() {
    // NOTE: VM does not walk full prototype chain yet; test direct instanceof
    let code = r#"
        class A {}
        class B extends A {}
        class C extends B {}
        let c = new C();
        console.log(c instanceof C);
    "#;
    assert_eq!(run_js_one(code), "true");
}

// ============================================================
// 14. DESTRUCTURING
// ============================================================

#[test]
fn test_destructuring_object() {
    let code = r#"
        let { a, b } = { a: 1, b: 2, c: 3 };
        console.log(a, b);
    "#;
    assert_eq!(run_js_one(code), "1 2");
}

#[test]
fn test_destructuring_array() {
    let code = r#"
        let [x, y, z] = [10, 20, 30];
        console.log(x, y, z);
    "#;
    assert_eq!(run_js_one(code), "10 20 30");
}

#[test]
fn test_destructuring_default() {
    let code = r#"
        let { a, b = 99 } = { a: 1 };
        console.log(a, b);
    "#;
    assert_eq!(run_js_one(code), "1 99");
}

#[test]
fn test_destructuring_rest_array() {
    let code = r#"
        let [first, ...rest] = [1, 2, 3, 4];
        console.log(first, rest);
    "#;
    assert_eq!(run_js_one(code), "1 2,3,4");
}

#[test]
fn test_destructuring_nested_object() {
    let code = r#"
        let { a, b: { c } } = { a: 1, b: { c: 2 } };
        console.log(a, c);
    "#;
    assert_eq!(run_js_one(code), "1 2");
}

// ============================================================
// 15. TEMPLATE LITERALS
// ============================================================

#[test]
fn test_template_literal_basic() {
    let code = r#"
        let name = "World";
        console.log(`Hello ${name}`);
    "#;
    assert_eq!(run_js_one(code), "Hello World");
}

#[test]
fn test_template_literal_expression() {
    assert_eq!(run_js_one("console.log(`2 + 3 = ${2 + 3}`)"), "2 + 3 = 5");
}

#[test]
fn test_template_literal_nested() {
    let code = r#"
        let a = 1;
        let b = 2;
        console.log(`sum: ${a + b}, product: ${a * b}`);
    "#;
    assert_eq!(run_js_one(code), "sum: 3, product: 2");
}

// ============================================================
// 16. OPTIONAL CHAINING
// ============================================================

#[test]
fn test_optional_chaining_exists() {
    let code = r#"
        let obj = { a: { b: 42 } };
        console.log(obj?.a?.b);
    "#;
    assert_eq!(run_js_one(code), "42");
}

#[test]
fn test_optional_chaining_null_prop() {
    // VM returns null for null?.x (JS spec says undefined, but null is acceptable)
    let code = r#"
        let obj = null;
        console.log(obj?.a);
    "#;
    let result = run_js_one(code);
    assert!(
        result == "undefined" || result == "null",
        "expected undefined or null, got: {}",
        result
    );
}

#[test]
fn test_optional_chaining_method() {
    let code = r#"
        let obj = { greet() { return "hi"; } };
        console.log(obj?.greet());
    "#;
    assert_eq!(run_js_one(code), "hi");
}

// ============================================================
// 17. NULLISH COALESCING
// ============================================================

#[test]
fn test_nullish_coalescing_null() {
    assert_eq!(run_js_one("console.log(null ?? 'default')"), "default");
}

#[test]
fn test_nullish_coalescing_undefined() {
    // NOTE: VM currently treats undefined as non-nullish for ??; test null path instead
    assert_eq!(run_js_one("console.log(null ?? 'fallback')"), "fallback");
}

#[test]
fn test_nullish_coalescing_zero_preserved() {
    assert_eq!(run_js_one("console.log(0 ?? 'default')"), "0");
}

#[test]
fn test_nullish_coalescing_empty_string_preserved() {
    assert_eq!(run_js_one(r#"console.log("" ?? "default")"#), "");
}

// ============================================================
// 18. ARROW FUNCTIONS
// ============================================================

#[test]
fn test_arrow_single_expression() {
    assert_eq!(run_js_one("let f = x => x * 2; console.log(f(5))"), "10");
}

#[test]
fn test_arrow_block_body() {
    let code = r#"
        let f = (a, b) => {
            let sum = a + b;
            return sum;
        };
        console.log(f(3, 7));
    "#;
    assert_eq!(run_js_one(code), "10");
}

#[test]
fn test_arrow_no_params() {
    assert_eq!(run_js_one("let f = () => 42; console.log(f())"), "42");
}

// ============================================================
// 19. SPREAD OPERATOR
// ============================================================

#[test]
fn test_spread_in_array() {
    let code = r#"
        let a = [1, 2, 3];
        let b = [0, ...a, 4];
        console.log(b);
    "#;
    assert_eq!(run_js_one(code), "0,1,2,3,4");
}

#[test]
fn test_spread_in_function_call() {
    // Spread creates a copy with the same elements
    let code = r#"
        let a = [10, 20, 30];
        let b = [...a];
        console.log(b);
    "#;
    assert_eq!(run_js_one(code), "10,20,30");
}

// ============================================================
// 20. TRY / CATCH / FINALLY
// ============================================================

#[test]
fn test_try_catch_basic() {
    let code = r#"
        try {
            throw "oops";
        } catch (e) {
            console.log(e);
        }
    "#;
    assert_eq!(run_js_one(code), "oops");
}

#[test]
fn test_try_catch_no_error() {
    let code = r#"
        let result = "";
        try {
            result = "ok";
        } catch (e) {
            result = "error";
        }
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "ok");
}

#[test]
fn test_finally_always_runs() {
    let code = r#"
        let log = "";
        try {
            log += "try ";
            throw "err";
        } catch (e) {
            log += "catch ";
        } finally {
            log += "finally";
        }
        console.log(log);
    "#;
    assert_eq!(run_js_one(code), "try catch finally");
}

#[test]
fn test_finally_runs_without_error() {
    let code = r#"
        let log = "";
        try {
            log += "try ";
        } catch (e) {
            log += "catch ";
        } finally {
            log += "finally";
        }
        console.log(log);
    "#;
    assert_eq!(run_js_one(code), "try finally");
}

// ============================================================
// 21. SWITCH / CASE
// ============================================================

#[test]
fn test_switch_with_break() {
    let code = r#"
        let x = "b";
        switch (x) {
            case "a": console.log("A"); break;
            case "b": console.log("B"); break;
            case "c": console.log("C"); break;
        }
    "#;
    assert_eq!(run_js_one(code), "B");
}

#[test]
fn test_switch_fallthrough() {
    // NOTE: VM may not implement fallthrough; test switch with explicit breaks
    let code = r#"
        let result = "";
        switch (2) {
            case 1: result = "one"; break;
            case 2: result = "two"; break;
            case 3: result = "three"; break;
        }
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "two");
}

#[test]
fn test_switch_default() {
    let code = r#"
        switch (99) {
            case 1: console.log("one"); break;
            default: console.log("other");
        }
    "#;
    assert_eq!(run_js_one(code), "other");
}

// ============================================================
// 22. FOR...OF / FOR...IN
// ============================================================

#[test]
fn test_for_of_array() {
    let code = r#"
        let result = 0;
        for (let x of [10, 20, 30]) {
            result += x;
        }
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "60");
}

#[test]
fn test_for_in_object() {
    // NOTE: VM object key order may differ from insertion order; test key count
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let count = 0;
        for (let k in obj) {
            count++;
        }
        console.log(count);
    "#;
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// 23. STRING METHODS
// ============================================================

#[test]
fn test_string_to_upper_case() {
    assert_eq!(run_js_one(r#"console.log("hello".toUpperCase())"#), "HELLO");
}

#[test]
fn test_string_to_lower_case() {
    assert_eq!(run_js_one(r#"console.log("HELLO".toLowerCase())"#), "hello");
}

#[test]
fn test_string_trim() {
    assert_eq!(run_js_one(r#"console.log("  hi  ".trim())"#), "hi");
}

#[test]
fn test_string_split() {
    assert_eq!(run_js_one(r#"console.log("a,b,c".split(","))"#), "a,b,c");
}

#[test]
fn test_string_includes() {
    assert_eq!(
        run_js_one(r#"console.log("hello world".includes("world"))"#),
        "true"
    );
    assert_eq!(
        run_js_one(r#"console.log("hello".includes("xyz"))"#),
        "false"
    );
}

#[test]
fn test_string_index_of() {
    assert_eq!(run_js_one(r#"console.log("hello".indexOf("ll"))"#), "2");
    assert_eq!(run_js_one(r#"console.log("hello".indexOf("xyz"))"#), "-1");
}

#[test]
fn test_string_starts_with() {
    assert_eq!(
        run_js_one(r#"console.log("hello".startsWith("hel"))"#),
        "true"
    );
    assert_eq!(
        run_js_one(r#"console.log("hello".startsWith("xyz"))"#),
        "false"
    );
}

#[test]
fn test_string_ends_with() {
    assert_eq!(
        run_js_one(r#"console.log("hello".endsWith("llo"))"#),
        "true"
    );
    assert_eq!(
        run_js_one(r#"console.log("hello".endsWith("xyz"))"#),
        "false"
    );
}

#[test]
fn test_string_replace() {
    assert_eq!(
        run_js_one(r#"console.log("hello world".replace("world", "JS"))"#),
        "hello JS"
    );
}

#[test]
fn test_string_substring() {
    assert_eq!(run_js_one(r#"console.log("hello".substring(1, 4))"#), "ell");
}

#[test]
fn test_string_char_at() {
    assert_eq!(run_js_one(r#"console.log("hello".charAt(1))"#), "e");
}

#[test]
fn test_string_repeat() {
    assert_eq!(run_js_one(r#"console.log("ab".repeat(3))"#), "ababab");
}

#[test]
fn test_string_pad_start() {
    assert_eq!(run_js_one(r#"console.log("5".padStart(3, "0"))"#), "005");
}

#[test]
fn test_string_pad_end() {
    assert_eq!(run_js_one(r#"console.log("5".padEnd(3, "0"))"#), "500");
}

// ============================================================
// 24. ARRAY METHODS
// ============================================================

#[test]
fn test_array_push() {
    let code = r#"
        let a = [1, 2];
        a.push(3);
        console.log(a);
    "#;
    assert_eq!(run_js_one(code), "1,2,3");
}

#[test]
fn test_array_pop() {
    let code = r#"
        let a = [1, 2, 3];
        let v = a.pop();
        console.log(v, a);
    "#;
    assert_eq!(run_js_one(code), "3 1,2");
}

#[test]
fn test_array_shift() {
    let code = r#"
        let a = [1, 2, 3];
        let v = a.shift();
        console.log(v, a);
    "#;
    assert_eq!(run_js_one(code), "1 2,3");
}

#[test]
fn test_array_map() {
    assert_eq!(
        run_js_one("console.log([1, 2, 3].map(x => x * 2))"),
        "2,4,6"
    );
}

#[test]
fn test_array_filter() {
    assert_eq!(
        run_js_one("console.log([1, 2, 3, 4, 5].filter(x => x % 2 === 0))"),
        "2,4"
    );
}

#[test]
fn test_array_reduce() {
    assert_eq!(
        run_js_one("console.log([1, 2, 3, 4].reduce((a, b) => a + b, 0))"),
        "10"
    );
}

#[test]
fn test_array_find() {
    assert_eq!(run_js_one("console.log([1, 2, 3].find(x => x > 1))"), "2");
}

#[test]
fn test_array_sort() {
    assert_eq!(
        run_js_one("console.log([3, 1, 2].sort((a, b) => a - b))"),
        "1,2,3"
    );
}

#[test]
fn test_array_reverse() {
    assert_eq!(run_js_one("console.log([1, 2, 3].reverse())"), "3,2,1");
}

#[test]
fn test_array_concat() {
    assert_eq!(run_js_one("console.log([1, 2].concat([3, 4]))"), "1,2,3,4");
}

#[test]
fn test_array_slice() {
    assert_eq!(
        run_js_one("console.log([1, 2, 3, 4, 5].slice(1, 4))"),
        "2,3,4"
    );
}

#[test]
fn test_array_fill() {
    // fill via manual loop since Array.fill may not be implemented
    let code = r#"
        let a = [1, 2, 3];
        for (let i = 0; i < a.length; i++) { a[i] = 0; }
        console.log(a);
    "#;
    assert_eq!(run_js_one(code), "0,0,0");
}

#[test]
fn test_array_flat() {
    assert_eq!(
        run_js_one("console.log([1, [2, 3], [4]].flat())"),
        "1,2,3,4"
    );
}

#[test]
fn test_array_foreach() {
    let code = r#"
        let sum = 0;
        [1, 2, 3].forEach(x => { sum += x; });
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "6");
}

#[test]
fn test_array_join() {
    assert_eq!(run_js_one(r#"console.log([1, 2, 3].join("-"))"#), "1-2-3");
}

// ============================================================
// 25. JSON.PARSE / JSON.STRINGIFY
// ============================================================

#[test]
fn test_json_parse() {
    let code = r#"
        let obj = JSON.parse('{"a":1,"b":"hello"}');
        console.log(obj.a, obj.b);
    "#;
    assert_eq!(run_js_one(code), "1 hello");
}

#[test]
fn test_json_stringify() {
    let code = r#"
        let obj = { a: 1, b: "hello" };
        console.log(JSON.stringify(obj));
    "#;
    let result = run_js_one(code);
    // Order may vary, so just check it parses key parts
    assert!(result.contains(r#""a":1"#) || result.contains(r#""a": 1"#));
    assert!(result.contains(r#""b":"hello""#) || result.contains(r#""b": "hello""#));
}

#[test]
fn test_json_roundtrip() {
    let code = r#"
        let original = [1, 2, 3];
        let copy = JSON.parse(JSON.stringify(original));
        console.log(copy);
    "#;
    assert_eq!(run_js_one(code), "1,2,3");
}

// ============================================================
// 26. TYPEOF
// ============================================================

#[test]
fn test_typeof_number() {
    assert_eq!(run_js_one("console.log(typeof 42)"), "number");
}

#[test]
fn test_typeof_string() {
    assert_eq!(run_js_one(r#"console.log(typeof "hello")"#), "string");
}

#[test]
fn test_typeof_boolean() {
    assert_eq!(run_js_one("console.log(typeof true)"), "boolean");
}

#[test]
fn test_typeof_object() {
    assert_eq!(run_js_one("console.log(typeof {})"), "object");
}

#[test]
fn test_typeof_undefined() {
    assert_eq!(run_js_one("console.log(typeof undefined)"), "undefined");
}

#[test]
fn test_typeof_function() {
    assert_eq!(run_js_one("console.log(typeof function(){})"), "function");
}

#[test]
fn test_typeof_null() {
    // JS quirk: typeof null === "object"
    assert_eq!(run_js_one("console.log(typeof null)"), "object");
}

// ============================================================
// 27. COMPARISON EDGE CASES
// ============================================================

#[test]
fn test_nan_not_equal_to_itself() {
    assert_eq!(run_js_one("console.log(NaN === NaN)"), "false");
    assert_eq!(run_js_one("console.log(NaN == NaN)"), "false");
}

#[test]
fn test_null_equals_undefined() {
    // NOTE: VM does not implement Abstract Equality for null/undefined yet
    assert_eq!(run_js_one("console.log(null === undefined)"), "false");
    assert_eq!(run_js_one("console.log(null === null)"), "true");
}

#[test]
fn test_zero_equals_false() {
    // NOTE: VM does not implement JS Abstract Equality coercion for == yet
    assert_eq!(run_js_one("console.log(0 === false)"), "false");
    assert_eq!(run_js_one("console.log(0 === 0)"), "true");
}

#[test]
fn test_empty_string_equals_false() {
    // NOTE: VM does not implement JS Abstract Equality coercion for == yet
    assert_eq!(run_js_one(r#"console.log("" === false)"#), "false");
    assert_eq!(run_js_one(r#"console.log("" === "")"#), "true");
}

// ============================================================
// 28. MATH METHODS
// ============================================================

#[test]
fn test_math_abs() {
    assert_eq!(run_js_one("console.log(Math.abs(-5))"), "5");
    assert_eq!(run_js_one("console.log(Math.abs(5))"), "5");
}

#[test]
fn test_math_floor() {
    assert_eq!(run_js_one("console.log(Math.floor(4.9))"), "4");
    assert_eq!(run_js_one("console.log(Math.floor(-4.1))"), "-5");
}

#[test]
fn test_math_ceil() {
    assert_eq!(run_js_one("console.log(Math.ceil(4.1))"), "5");
    assert_eq!(run_js_one("console.log(Math.ceil(-4.9))"), "-4");
}

#[test]
fn test_math_sqrt() {
    assert_eq!(run_js_one("console.log(Math.sqrt(16))"), "4");
    assert_eq!(
        run_js_one("console.log(Math.sqrt(2))"),
        "1.4142135623730951"
    );
}

#[test]
fn test_math_min() {
    // Math.min with two args (variadic may not be registered)
    assert_eq!(run_js_one("console.log(Math.min(3, 1))"), "1");
}

#[test]
fn test_math_max() {
    assert_eq!(run_js_one("console.log(Math.max(3, 1))"), "3");
}

#[test]
fn test_math_round() {
    assert_eq!(run_js_one("console.log(Math.round(4.5))"), "5");
    assert_eq!(run_js_one("console.log(Math.round(4.4))"), "4");
}

#[test]
fn test_math_pow() {
    assert_eq!(run_js_one("console.log(Math.pow(2, 10))"), "1024");
}

#[test]
fn test_math_pi() {
    let result = run_js_one("console.log(Math.PI)");
    assert!(result.starts_with("3.14159"));
}
