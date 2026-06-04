use super::helpers::{run_js, run_js_vm};
/// Tests for JS → host function interop: objects crossing the boundary,
/// namespace resolution, Map/Set via host, invoke from Rust.
use std::sync::Arc;
use vybe_bytecode::{VM, Value};

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

// ============================================================
// A. MAP HOST OBJECT
// ============================================================

#[test]
fn map_set_get_has_delete_size() {
    let out = run_js(
        r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        console.log(m.size);
        console.log(m.get("a"));
        console.log(m.has("b"));
        m.delete("a");
        console.log(m.size);
        console.log(m.has("a"));
    "#,
    );
    assert_eq!(out, vec!["2", "1", "true", "1", "false"]);
}

#[test]
fn map_clear() {
    let out = run_js(
        r#"
        let m = new Map();
        m.set("x", 10);
        m.set("y", 20);
        m.clear();
        console.log(m.size);
    "#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn map_keys_values() {
    let out = run_js(
        r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        let k = m.keys();
        let v = m.values();
        console.log(k.length);
        console.log(v.length);
    "#,
    );
    assert_eq!(out, vec!["2", "2"]);
}

// ============================================================
// B. SET HOST OBJECT
// ============================================================

#[test]
fn set_add_has_delete_size() {
    let out = run_js(
        r#"
        let s = new Set();
        s.add(1);
        s.add(2);
        s.add(2);
        console.log(s.size);
        console.log(s.has(1));
        s.delete(1);
        console.log(s.size);
        console.log(s.has(1));
    "#,
    );
    assert_eq!(out, vec!["2", "true", "1", "false"]);
}

#[test]
fn set_clear() {
    let out = run_js(
        r#"
        let s = new Set();
        s.add("a");
        s.add("b");
        s.clear();
        console.log(s.size);
    "#,
    );
    assert_eq!(out, vec!["0"]);
}

// ============================================================
// C. JSON HOST FUNCTIONS
// ============================================================

#[test]
fn json_parse_object() {
    let out = run_js(
        r#"
        let obj = JSON.parse('{"name":"test","age":25}');
        console.log(obj.name);
        console.log(obj.age);
    "#,
    );
    assert_eq!(out, vec!["test", "25"]);
}

#[test]
fn json_parse_array() {
    let out = run_js(
        r#"
        let arr = JSON.parse('[1,2,3]');
        console.log(arr.length);
    "#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn json_stringify() {
    let out = run_js(
        r#"
        let s = JSON.stringify({x: 1});
        console.log(typeof s);
    "#,
    );
    assert_eq!(out, vec!["string"]);
}

// ============================================================
// D. MATH NAMESPACE
// ============================================================

#[test]
fn math_abs() {
    assert_eq!(run_js_one("console.log(Math.abs(-5))"), "5");
}

#[test]
fn math_floor() {
    assert_eq!(run_js_one("console.log(Math.floor(3.7))"), "3");
}

#[test]
fn math_ceil() {
    assert_eq!(run_js_one("console.log(Math.ceil(3.2))"), "4");
}

#[test]
fn math_sqrt() {
    assert_eq!(run_js_one("console.log(Math.sqrt(16))"), "4");
}

#[test]
fn math_round() {
    assert_eq!(run_js_one("console.log(Math.round(3.5))"), "4");
}

#[test]
fn math_min_max() {
    let out = run_js(
        r#"
        console.log(Math.min(3, 1));
        console.log(Math.max(1, 3));
    "#,
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn math_pow() {
    assert_eq!(run_js_one("console.log(Math.pow(2, 10))"), "1024");
}

#[test]
fn math_pi() {
    let out = run_js_one("console.log(Math.PI > 3.14)");
    assert_eq!(out, "true");
}

// ============================================================
// E. OBJECT STATIC METHODS → HOST
// ============================================================

#[test]
fn object_keys() {
    let out = run_js(
        r#"
        let obj = {a: 1, b: 2, c: 3};
        console.log(Object.keys(obj).length);
    "#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn object_values() {
    let out = run_js(
        r#"
        let obj = {x: 10};
        let v = Object.values(obj);
        console.log(v.length);
    "#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn object_entries() {
    let out = run_js(
        r#"
        let obj = {a: 1, b: 2};
        console.log(Object.entries(obj).length);
    "#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn object_assign() {
    let out = run_js(
        r#"
        let a = {x: 1};
        let b = {y: 2};
        Object.assign(a, b);
        console.log(a.x);
        console.log(a.y);
    "#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn object_from_entries() {
    let out = run_js(
        r#"
        let entries = [["a", 1], ["b", 2]];
        let obj = Object.fromEntries(entries);
        console.log(obj.a);
        console.log(obj.b);
    "#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

// ============================================================
// F. ARRAY.FROM / ARRAY.ISARRAY → HOST
// ============================================================

#[test]
fn array_from_copies() {
    let out = run_js(
        r#"
        let orig = [1, 2, 3];
        let copy = Array.from(orig);
        copy.push(4);
        console.log(orig.length);
        console.log(copy.length);
    "#,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn array_is_array() {
    let out = run_js(
        r#"
        console.log(Array.isArray([1,2]));
        console.log(Array.isArray("hello"));
        console.log(Array.isArray(42));
    "#,
    );
    assert_eq!(out, vec!["true", "false", "false"]);
}

// ============================================================
// G. INVOKE JS FUNCTION FROM RUST
// ============================================================

#[test]
fn invoke_js_global_function() {
    let (mut vm, output) = run_js_vm(
        r#"
        function greet(name) {
            console.log("hello " + name);
        }
    "#,
    );
    let func = vm.globals.get("greet").cloned().unwrap();
    vm.invoke(&func, &[Value::String(Arc::from("world"))])
        .unwrap();
    assert_eq!(
        output.lock().unwrap().last().map(|s| s.as_str()),
        Some("hello world")
    );
}

#[test]
fn invoke_js_class_method() {
    let (mut vm, output) = run_js_vm(
        r#"
        class Counter {
            constructor() { this.n = 0; }
            inc() { this.n = this.n + 1; }
            report() { console.log(this.n); }
        }
        var c = new Counter();
    "#,
    );
    let instance = vm.globals.get("c").cloned();
    assert!(instance.is_some(), "var c should be a global");
    let instance = instance.unwrap();
    let inc = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("inc").cloned()
    } else {
        None
    }
    .unwrap();
    let report = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("report").cloned()
    } else {
        None
    }
    .unwrap();

    // Host-driven JS method invocation needs `__js_this` bound to the
    // receiver before each call — the VM's `invoke` helper isn't
    // method-aware (mirrors `Function.prototype.call(receiver, …)`'s
    // explicit-receiver convention, which the test mimics by passing
    // `instance` as the first arg). Without the bind, the method
    // body's `this` reads stale `__js_this` and `this.n` traps.
    let bind_this = |vm: &mut VM, recv: &Value| {
        vm.globals.insert("__js_this".to_string(), recv.clone());
    };
    bind_this(&mut vm, &instance);
    vm.invoke(&inc, &[instance.clone()]).unwrap();
    bind_this(&mut vm, &instance);
    vm.invoke(&inc, &[instance.clone()]).unwrap();
    bind_this(&mut vm, &instance);
    vm.invoke(&report, &[instance.clone()]).unwrap();
    assert_eq!(output.lock().unwrap().last().map(|s| s.as_str()), Some("2"));
}

#[test]
fn invoke_preserves_state() {
    let (mut vm, output) = run_js_vm(
        r#"
        let count = 0;
        function inc() { count++; console.log(count); }
    "#,
    );
    let func = vm.globals.get("inc").cloned().unwrap();
    vm.invoke(&func, &[]).unwrap();
    vm.invoke(&func, &[]).unwrap();
    vm.invoke(&func, &[]).unwrap();
    let out = output.lock().unwrap();
    assert_eq!(out.as_slice(), &["1", "2", "3"]);
}

// ============================================================
// H. IN / DELETE / HASOWNPROPERTY → HOST
// ============================================================

#[test]
fn in_operator_calls_host() {
    let out = run_js(
        r#"
        let obj = {a: 1, b: 2};
        console.log("a" in obj);
        console.log("c" in obj);
    "#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn delete_operator_calls_host() {
    let out = run_js(
        r#"
        let obj = {a: 1, b: 2};
        delete obj.a;
        console.log("a" in obj);
        console.log("b" in obj);
    "#,
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn has_own_property_calls_host() {
    let out = run_js(
        r#"
        let obj = {x: 1};
        console.log(obj.hasOwnProperty("x"));
        console.log(obj.hasOwnProperty("y"));
    "#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

// ============================================================
// I. STRING METHODS (opcode intrinsics)
// ============================================================

#[test]
fn string_to_upper_lower() {
    let out = run_js(
        r#"
        console.log("hello".toUpperCase());
        console.log("WORLD".toLowerCase());
    "#,
    );
    assert_eq!(out, vec!["HELLO", "world"]);
}

#[test]
fn string_trim() {
    assert_eq!(run_js_one(r#"console.log("  hi  ".trim())"#), "hi");
}

#[test]
fn string_split_join() {
    assert_eq!(
        run_js_one(r#"console.log("a,b,c".split(",").join("-"))"#),
        "a-b-c"
    );
}

#[test]
fn string_starts_ends_with() {
    let out = run_js(
        r#"
        console.log("hello".startsWith("hel"));
        console.log("hello".endsWith("llo"));
        console.log("hello".startsWith("xyz"));
    "#,
    );
    assert_eq!(out, vec!["true", "true", "false"]);
}

#[test]
fn string_index_of() {
    let out = run_js(
        r#"
        console.log("hello world".indexOf("world"));
        console.log("hello".indexOf("xyz"));
    "#,
    );
    assert_eq!(out, vec!["6", "-1"]);
}

#[test]
fn string_substring_replace() {
    let out = run_js(
        r#"
        console.log("hello world".substring(6, 11));
        console.log("hello world".replace("world", "js"));
    "#,
    );
    assert_eq!(out, vec!["world", "hello js"]);
}

#[test]
fn string_char_at() {
    assert_eq!(run_js_one(r#"console.log("hello".charAt(1))"#), "e");
}

#[test]
fn string_repeat_pad() {
    let out = run_js(
        r#"
        console.log("ab".repeat(3));
        console.log("5".padStart(3, "0"));
    "#,
    );
    assert_eq!(out, vec!["ababab", "005"]);
}

// ============================================================
// J. ARRAY METHODS (opcode intrinsics + host)
// ============================================================

#[test]
fn array_push_pop_length() {
    let out = run_js(
        r#"
        let a = [1, 2];
        a.push(3);
        console.log(a.length);
        let x = a.pop();
        console.log(x);
        console.log(a.length);
    "#,
    );
    assert_eq!(out, vec!["3", "3", "2"]);
}

#[test]
fn array_map_filter() {
    let out = run_js(
        r#"
        let a = [1, 2, 3, 4, 5];
        let evens = a.filter(x => x % 2 === 0);
        let doubled = evens.map(x => x * 2);
        console.log(doubled.join(","));
    "#,
    );
    assert_eq!(out, vec!["4,8"]);
}

#[test]
fn array_reduce() {
    assert_eq!(
        run_js_one(r#"console.log([1,2,3,4].reduce((a, b) => a + b, 0))"#),
        "10"
    );
}

#[test]
fn array_find() {
    assert_eq!(
        run_js_one(r#"console.log([10, 20, 30].find(x => x > 15))"#),
        "20"
    );
}

#[test]
fn array_some_every() {
    let out = run_js(
        r#"
        console.log([1,2,3].some(x => x > 2));
        console.log([1,2,3].every(x => x > 0));
        console.log([1,2,3].every(x => x > 1));
    "#,
    );
    assert_eq!(out, vec!["true", "true", "false"]);
}

#[test]
fn array_sort_reverse_concat() {
    let out = run_js(
        r#"
        let a = [3, 1, 2];
        a.sort((a, b) => a - b);
        console.log(a.join(","));
        a.reverse();
        console.log(a.join(","));
        let b = a.concat([4, 5]);
        console.log(b.join(","));
    "#,
    );
    assert_eq!(out, vec!["1,2,3", "3,2,1", "3,2,1,4,5"]);
}

#[test]
fn array_slice() {
    assert_eq!(
        run_js_one(r#"console.log([1,2,3,4,5].slice(1,3).join(","))"#),
        "2,3"
    );
}
