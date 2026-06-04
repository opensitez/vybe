/// Optional chaining edge cases — optional call, delete with ?., short-circuit
/// with side effects, null vs undefined, optional index, chained ?.
use super::helpers::run_js;

// ── optional call ─────────────────────────────────────────────────────────────

#[test]
fn optional_call_invokes_if_function() {
    assert_eq!(
        run_js(
            r#"
const fn1 = () => "called";
console.log(fn1?.());
"#
        ),
        vec!["called"]
    );
}

#[test]
fn optional_call_returns_undefined_if_not_function() {
    assert_eq!(
        run_js(
            r#"
const notFn = null;
console.log(notFn?.());
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn optional_call_does_not_throw_if_undefined() {
    assert_eq!(
        run_js(
            r#"
let x;
console.log(x?.());
"#
        ),
        vec!["undefined"]
    );
}

// ── side-effect short circuit ─────────────────────────────────────────────────

#[test]
fn optional_chain_short_circuits_member_access() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
function sideEffect() { count++; return 1; }
const obj = null;
obj?.prop[sideEffect()]; // sideEffect should NOT be called
console.log(count);
"#
        ),
        vec!["0"]
    );
}

#[test]
fn optional_chain_short_circuits_function_call() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
const obj = null;
obj?.method(count++); // count++ should NOT run
console.log(count);
"#
        ),
        vec!["0"]
    );
}

// ── optional computed access ──────────────────────────────────────────────────

#[test]
fn optional_computed_access_on_array() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(arr?.[1]);
const nothing = null;
console.log(nothing?.[0]);
"#
        ),
        vec!["2", "undefined"]
    );
}

#[test]
fn optional_dynamic_key_access() {
    assert_eq!(
        run_js(
            r#"
const key = "name";
const obj = { name: "Alice" };
console.log(obj?.[key]);
const absent = undefined;
console.log(absent?.[key]);
"#
        ),
        vec!["Alice", "undefined"]
    );
}

// ── chained optional ──────────────────────────────────────────────────────────

#[test]
fn chained_optional_access_deep() {
    assert_eq!(
        run_js(
            r#"
const a = { b: { c: { d: 42 } } };
console.log(a?.b?.c?.d);
console.log(a?.b?.x?.d);
console.log(a?.z?.c?.d);
"#
        ),
        vec!["42", "undefined", "undefined"]
    );
}

#[test]
fn mixed_optional_and_required_access() {
    assert_eq!(
        run_js(
            r#"
const config = {
    server: { host: "localhost", port: 8080 }
};
// Only the first access is optional
console.log(config?.server.host);
console.log(config?.missing?.port);
"#
        ),
        vec!["localhost", "undefined"]
    );
}

// ── null vs undefined ─────────────────────────────────────────────────────────

#[test]
fn optional_chain_triggers_on_both_null_and_undefined() {
    assert_eq!(
        run_js(
            r#"
const a = null;
const b = undefined;
console.log(a?.x);
console.log(b?.x);
"#
        ),
        vec!["undefined", "undefined"]
    );
}

#[test]
fn optional_chain_does_not_trigger_on_zero_or_false() {
    assert_eq!(
        run_js(
            r#"
const zero = 0;
const bool = false;
// 0?.x is different — 0 has no property 'x', but no short-circuit either
console.log(typeof zero?.toString);
console.log(typeof bool?.toString);
"#
        ),
        vec!["function", "function"]
    );
}

// ── optional with default via ?? ──────────────────────────────────────────────

#[test]
fn optional_chain_with_nullish_coalescing() {
    assert_eq!(
        run_js(
            r#"
const user = null;
const name = user?.profile?.name ?? "Guest";
console.log(name);
"#
        ),
        vec!["Guest"]
    );
}

#[test]
fn optional_chain_nullish_fallback_on_undefined() {
    assert_eq!(
        run_js(
            r#"
const cfg = {};
const port = cfg.server?.port ?? 3000;
console.log(port);
"#
        ),
        vec!["3000"]
    );
}

// ── optional method call on prototype ────────────────────────────────────────

#[test]
fn optional_method_on_string() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s?.toUpperCase());
const n = null;
console.log(n?.toUpperCase());
"#
        ),
        vec!["HELLO", "undefined"]
    );
}

// ── delete with optional chain ─────────────────────────────────────────────────

#[test]
fn delete_optional_chain_on_null_is_safe() {
    assert_eq!(
        run_js(
            r#"
const obj = null;
// delete obj?.prop should not throw when obj is null
const result = delete obj?.prop;
console.log(result);
"#
        ),
        vec!["true"]
    );
}

// ── optional in assignment context ────────────────────────────────────────────

#[test]
fn optional_chain_not_valid_on_left_of_assignment() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    eval("const x = {}; x?.y = 1;");
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── optional in template literals ────────────────────────────────────────────

#[test]
fn optional_chain_result_in_template() {
    assert_eq!(
        run_js(
            r#"
const user = { name: "Bob" };
const absent = null;
console.log(`Hello ${user?.name}`);
console.log(`Hello ${absent?.name ?? "stranger"}`);
"#
        ),
        vec!["Hello Bob", "Hello stranger"]
    );
}
