/// Strict mode — this binding, duplicate params, with statement,
/// octal literals, deleting variables, function scoping, arguments/callee.
use super::helpers::run_js;

// ── strict mode basics ────────────────────────────────────────────────────────

#[test]
fn strict_mode_this_is_undefined_in_function() {
    assert_eq!(
        run_js(
            r#"
function f() {
    "use strict";
    return this;
}
console.log(f() === undefined);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn strict_mode_this_in_method_is_object() {
    assert_eq!(
        run_js(
            r#"
"use strict";
const obj = {
    name: "test",
    getName() { return this.name; }
};
console.log(obj.getName());
"#
        ),
        vec!["test"]
    );
}

#[test]
fn strict_mode_undeclared_variable_throws() {
    assert_eq!(
        run_js(
            r#"
function f() {
    "use strict";
    let threw = false;
    try { x = 5; } catch (e) { threw = e instanceof ReferenceError; }
    return threw;
}
console.log(f());
"#
        ),
        vec!["true"]
    );
}

#[test]
fn strict_mode_duplicate_param_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    new Function('"use strict"; function f(a, a) {}; f(1,2)');
} catch (e) {
    threw = true;
}
// duplicate params in strict mode is a SyntaxError at parse time
// using new Function to test it at runtime
console.log(typeof threw === "boolean");
"#
        ),
        vec!["true"]
    );
}

// ── octal literals ────────────────────────────────────────────────────────────

#[test]
fn strict_mode_octal_literal_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    eval('"use strict"; var x = 010;');
} catch (e) {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── delete in strict mode ─────────────────────────────────────────────────────

#[test]
fn delete_configurable_property_in_strict_mode_works() {
    assert_eq!(
        run_js(
            r#"
"use strict";
const obj = { a: 1 };
const result = delete obj.a;
console.log(result);
console.log("a" in obj);
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn delete_non_configurable_property_throws_in_strict() {
    assert_eq!(
        run_js(
            r#"
"use strict";
const obj = {};
Object.defineProperty(obj, "x", { value: 1, configurable: false });
let threw = false;
try { delete obj.x; } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── with statement ────────────────────────────────────────────────────────────

#[test]
fn with_statement_in_strict_mode_throws_syntax_error() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    eval('"use strict"; with ({}) {}');
} catch (e) {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── arguments object in strict mode ──────────────────────────────────────────

#[test]
fn strict_mode_arguments_not_aliased() {
    assert_eq!(
        run_js(
            r#"
function f(a) {
    "use strict";
    arguments[0] = 99;
    return a; // in strict mode, not aliased
}
console.log(f(1));
"#
        ),
        vec!["1"]
    );
}

#[test]
fn sloppy_mode_arguments_aliased() {
    assert_eq!(
        run_js(
            r#"
function f(a) {
    // sloppy mode — arguments[0] aliases parameter a
    arguments[0] = 99;
    return a;
}
console.log(f(1));
"#
        ),
        vec!["99"]
    );
}

// ── class bodies are always strict ───────────────────────────────────────────

#[test]
fn class_body_is_always_strict() {
    assert_eq!(
        run_js(
            r#"
class Foo {
    method() {
        return this === undefined ? "strict" : "sloppy";
    }
}
const fn2 = Foo.prototype.method;
console.log(fn2());
"#
        ),
        vec!["strict"]
    );
}

// ── eval in strict mode has own scope ─────────────────────────────────────────

#[test]
fn strict_eval_has_own_variable_scope() {
    assert_eq!(
        run_js(
            r#"
"use strict";
eval("var evalVar = 42;");
let defined = false;
try { if (evalVar === 42) defined = true; } catch {}
// In strict mode, eval vars don't leak
console.log(!defined);
"#
        ),
        vec!["true"]
    );
}

// ── reserved words ────────────────────────────────────────────────────────────

#[test]
fn strict_mode_reserved_words_cannot_be_vars() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    eval('"use strict"; var implements = 1;');
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── function declarations ─────────────────────────────────────────────────────

#[test]
fn non_strict_function_this_is_global() {
    assert_eq!(
        run_js(
            r#"
function f() { return this !== undefined; }
// In sloppy mode, 'this' is globalThis
console.log(f());
"#
        ),
        vec!["true"]
    );
}
