/// Scope edge cases — let/const in blocks, var hoisting, TDZ in switch,
/// function declaration hoisting, for-loop let binding per-iteration,
/// closures capturing loop variable, IIFE scope, with-less patterns.

use super::helpers::run_js;

#[test]
fn let_per_iteration_in_for_loop() {
    assert_eq!(run_js(r#"
const fns = [];
for (let i = 0; i < 3; i++) {
    fns.push(() => i);
}
console.log(fns.map(f => f()).join(","));
"#), vec!["0,1,2"]);
}

#[test]
fn var_shared_across_loop_iterations() {
    assert_eq!(run_js(r#"
const fns = [];
for (var i = 0; i < 3; i++) {
    fns.push(() => i);
}
// All closures share same var i, which is 3 after loop
console.log(fns.map(f => f()).join(","));
"#), vec!["3,3,3"]);
}

#[test]
fn let_in_block_not_visible_outside() {
    assert_eq!(run_js(r#"
{
    let x = 42;
}
let threw = false;
try { x; } catch { threw = true; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn var_in_block_visible_outside() {
    assert_eq!(run_js(r#"
{
    var y = 99;
}
console.log(y);
"#), vec!["99"]);
}

#[test]
fn function_hoisted_before_var() {
    assert_eq!(run_js(r#"
console.log(fn()); // works — hoisted
function fn() { return "hoisted"; }
"#), vec!["hoisted"]);
}

#[test]
fn var_hoisted_but_undefined_before_init() {
    assert_eq!(run_js(r#"
console.log(typeof x); // undefined — hoisted but not initialized
var x = 42;
console.log(x);
"#), vec!["undefined", "42"]);
}

#[test]
fn let_tdz_throws_before_declaration() {
    assert_eq!(run_js(r#"
let threw = false;
try {
    eval("console.log(y); let y = 1;");
} catch (e) {
    threw = e instanceof ReferenceError;
}
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn const_not_reassignable() {
    assert_eq!(run_js(r#"
const x = 42;
let threw = false;
try { eval("const x = 42; x = 1;"); } catch { threw = true; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn const_object_mutable() {
    assert_eq!(run_js(r#"
const obj = { x: 1 };
obj.x = 99;
obj.y = 2;
console.log(obj.x);
console.log(obj.y);
"#), vec!["99", "2"]);
}

#[test]
fn iife_creates_scope() {
    assert_eq!(run_js(r#"
const result = (function() {
    const secret = 42;
    return secret;
})();
console.log(result);
let threw = false;
try { secret; } catch { threw = true; }
console.log(threw);
"#), vec!["42", "true"]);
}

#[test]
fn closure_captures_by_reference_not_value() {
    assert_eq!(run_js(r#"
let x = 1;
const get = () => x;
const set = v => { x = v; };
console.log(get());
set(42);
console.log(get());
"#), vec!["1", "42"]);
}

#[test]
fn switch_shares_let_binding_across_cases() {
    assert_eq!(run_js(r#"
switch (1) {
    case 1:
        let v = "from 1";
        // fall through
    case 2:
        // v visible here (same block)
        console.log(v);
        break;
}
"#), vec!["from 1"]);
}

#[test]
fn nested_function_scope_chain() {
    assert_eq!(run_js(r#"
const outer = 1;
function level1() {
    const mid = 2;
    function level2() {
        const inner = 3;
        return outer + mid + inner;
    }
    return level2();
}
console.log(level1());
"#), vec!["6"]);
}

#[test]
fn shadowing_with_inner_let() {
    assert_eq!(run_js(r#"
const x = "outer";
{
    const x = "inner";
    console.log(x);
}
console.log(x);
"#), vec!["inner", "outer"]);
}

#[test]
fn for_of_let_binding_per_iteration() {
    assert_eq!(run_js(r#"
const fns = [];
for (const v of [10, 20, 30]) {
    fns.push(() => v);
}
console.log(fns.map(f => f()).join(","));
"#), vec!["10,20,30"]);
}
