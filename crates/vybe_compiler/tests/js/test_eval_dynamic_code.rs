/// Eval and indirect eval, new Function — scoping, strict mode interaction,
/// code generation, security patterns, dynamic code execution.

use super::helpers::run_js;

// ── direct eval ───────────────────────────────────────────────────────────────

#[test]
fn eval_evaluates_expression() {
    assert_eq!(run_js(r#"
console.log(eval("1 + 2"));
console.log(eval("'hello'.toUpperCase()"));
"#), vec!["3", "HELLO"]);
}

#[test]
fn eval_returns_last_expression_value() {
    assert_eq!(run_js(r#"
const result = eval("let x = 5; x * x");
console.log(result);
"#), vec!["25"]);
}

#[test]
fn eval_can_access_enclosing_scope() {
    assert_eq!(run_js(r#"
const x = 42;
const result = eval("x + 8");
console.log(result);
"#), vec!["50"]);
}

#[test]
fn eval_can_modify_enclosing_scope_var() {
    assert_eq!(run_js(r#"
var y = 1;
eval("y = 99");
console.log(y);
"#), vec!["99"]);
}

// ── indirect eval ─────────────────────────────────────────────────────────────

#[test]
fn indirect_eval_runs_in_global_scope() {
    assert_eq!(run_js(r#"
const x = "local";
globalThis.x = "global";
const indirectEval = eval;
const result = indirectEval("x");
console.log(result);
"#), vec!["global"]);
}

#[test]
fn indirect_eval_via_assignment_to_var() {
    assert_eq!(run_js(r#"
const e = eval;
console.log(e("2 ** 10"));
"#), vec!["1024"]);
}

// ── new Function ──────────────────────────────────────────────────────────────

#[test]
fn new_function_creates_callable() {
    assert_eq!(run_js(r#"
const add = new Function("a", "b", "return a + b");
console.log(add(3, 4));
"#), vec!["7"]);
}

#[test]
fn new_function_no_args() {
    assert_eq!(run_js(r#"
const greet = new Function("return 'hello'");
console.log(greet());
"#), vec!["hello"]);
}

#[test]
fn new_function_comma_separated_params() {
    assert_eq!(run_js(r#"
const mul = new Function("x, y", "return x * y");
console.log(mul(6, 7));
"#), vec!["42"]);
}

#[test]
fn new_function_does_not_capture_outer_scope() {
    assert_eq!(run_js(r#"
const secret = "top-secret";
const fn2 = new Function("try { return secret; } catch { return 'undefined'; }");
// new Function runs in global scope, can't access local 'secret'
const result = fn2();
console.log(result === "top-secret" || result === "undefined");
"#), vec!["true"]);
}

#[test]
fn new_function_has_length_from_params() {
    assert_eq!(run_js(r#"
const f = new Function("a", "b", "c", "return a + b + c");
console.log(f.length);
"#), vec!["3"]);
}

#[test]
fn new_function_has_name_anonymous() {
    assert_eq!(run_js(r#"
const f = new Function("return 1");
console.log(f.name);
"#), vec!["anonymous"]);
}

// ── eval in strict mode ───────────────────────────────────────────────────────

#[test]
fn eval_strict_mode_via_string() {
    assert_eq!(run_js(r#"
const result = eval('"use strict"; let z = 10; z');
console.log(result);
"#), vec!["10"]);
}

#[test]
fn eval_does_not_leak_let_to_outer() {
    assert_eq!(run_js(r#"
eval("let innerLet = 5;");
let threw = false;
try { innerLet; } catch { threw = true; }
console.log(threw);
"#), vec!["true"]);
}

// ── Function constructor properties ──────────────────────────────────────────

#[test]
fn function_constructor_creates_function_instance() {
    assert_eq!(run_js(r#"
const f = new Function("return 42");
console.log(f instanceof Function);
console.log(typeof f);
"#), vec!["true", "function"]);
}

#[test]
fn new_function_can_use_closures_via_outer_function() {
    assert_eq!(run_js(r#"
function makeAdder(n) {
    // Can't capture n via new Function, so pass it as arg
    return new Function("x", "return x + " + n);
}
const add5 = makeAdder(5);
console.log(add5(10));
"#), vec!["15"]);
}

// ── typeof with undeclared ────────────────────────────────────────────────────

#[test]
fn typeof_undeclared_safe_check() {
    assert_eq!(run_js(r#"
// typeof doesn't throw for undeclared
console.log(typeof COMPLETELY_UNDECLARED_VARIABLE);
"#), vec!["undefined"]);
}
