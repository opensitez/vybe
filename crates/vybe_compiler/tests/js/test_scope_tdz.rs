/// Scope deep — TDZ (temporal dead zone), class TDZ, block-scoped functions,
/// const mutation, closure sharing, IIFE isolation, catch binding scope,
/// function param shadowing, let in switch, nested shadowing.
use super::helpers::run_js;

// ── TDZ — temporal dead zone ──────────────────────────────────────────────────

#[test]
fn const_tdz_in_block() {
    assert_eq!(
        run_js(
            r#"
let result;
{
    try {
        result = x;
        const x = 1;
    } catch (e) {
        result = "tdz";
    }
}
console.log(result);
"#
        ),
        vec!["tdz"]
    );
}

// ── nested scope shadowing ────────────────────────────────────────────────────

#[test]
fn inner_let_shadows_outer_let() {
    assert_eq!(
        run_js(
            r#"
let x = "outer";
{
    let x = "inner";
    console.log(x);
}
console.log(x);
"#
        ),
        vec!["inner", "outer"]
    );
}

#[test]
fn inner_var_overrides_outer_var_same_scope() {
    assert_eq!(
        run_js(
            r#"
var y = "first";
{
    var y = "second";
}
console.log(y);
"#
        ),
        vec!["second"]
    );
}

#[test]
fn function_param_shadows_outer_let() {
    assert_eq!(
        run_js(
            r#"
let value = "global";
function f(value) {
    return value;
}
console.log(f("local"));
console.log(value);
"#
        ),
        vec!["local", "global"]
    );
}

// ── class TDZ ─────────────────────────────────────────────────────────────────

#[test]
fn class_tdz_throws_before_declaration() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    new Foo();
    class Foo {}
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn class_declaration_not_hoisted_unlike_function() {
    // §9.1.1.1.6 GetBindingValue: `typeof Bar` on an uninitialized (TDZ)
    // class binding throws ReferenceError — typeof only shields
    // *unresolvable* names, not declared-but-uninitialized ones.
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    const x = typeof Bar;
    class Bar {}
    console.log(typeof new Bar());
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── block-scoped function declarations ───────────────────────────────────────

#[test]
fn function_in_strict_block_scoped_to_block() {
    assert_eq!(
        run_js(
            r#"
"use strict";
let result;
{
    function blockFn() { return 42; }
    result = blockFn();
}
console.log(result);
"#
        ),
        vec!["42"]
    );
}

// ── const must be initialized ─────────────────────────────────────────────────

#[test]
fn const_must_be_initialized_at_declaration() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    eval("const x;");
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── closure captures binding not value ───────────────────────────────────────

#[test]
fn closure_sees_updated_let_value() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
function increment() { count++; }
function get() { return count; }
increment();
increment();
console.log(get());
"#
        ),
        vec!["2"]
    );
}

#[test]
fn multiple_closures_share_same_binding() {
    assert_eq!(
        run_js(
            r#"
function makeCounter() {
    let n = 0;
    return {
        inc() { n++; },
        get() { return n; }
    };
}
const c = makeCounter();
c.inc(); c.inc(); c.inc();
console.log(c.get());
"#
        ),
        vec!["3"]
    );
}

// ── IIFE isolates scope ───────────────────────────────────────────────────────

#[test]
fn iife_prevents_var_leaking_to_outer() {
    assert_eq!(
        run_js(
            r#"
(function() {
    var localVar = "local";
})();
console.log(typeof localVar);
"#
        ),
        vec!["undefined"]
    );
}

// ── switch and scope ──────────────────────────────────────────────────────────

#[test]
fn let_in_switch_shared_across_cases() {
    assert_eq!(
        run_js(
            r#"
let result = "";
switch (1) {
    case 1:
        let x = "shared";
        result += x;
    case 2:
        result += "-done";
}
console.log(result);
"#
        ),
        vec!["shared-done"]
    );
}

// ── catch block binding ───────────────────────────────────────────────────────

#[test]
fn catch_binding_scoped_to_catch_block_var_escapes() {
    assert_eq!(
        run_js(
            r#"
try { throw new Error("test"); }
catch (e) { var caught = e.message; }
console.log(caught);
console.log(typeof e);
"#
        ),
        vec!["test", "undefined"]
    );
}
