//! Exception-coverage gaps not exercised elsewhere in the JS suite.
//!
//! The existing `test_error_*`, `test_ecma_error_handling`,
//! `test_control_flow_advanced` and `test_try_catch_finally_edge` files
//! already cover the eight Error constructors, the instanceof hierarchy,
//! `cause`/`AggregateError`, async rejection, and the `return`/`throw`
//! interactions with `finally`. This file fills the narrow remaining gaps:
//!
//!   1. Throwing/catching *non-Error primitive* values (§14.14: any
//!      ECMAScript value is throwable; `catch` binds it by identity).
//!   2. Error own/prototype property-descriptor details (§20.5).
//!   3. `finally` running when `break`/`continue` exits a loop `try`.

crate::js_cases! {
    // ── 1. Throwing every non-Error value type ──────────────────────────

    throw_null_is_caught_as_null => {
        r#"
let ok = false;
try { throw null; } catch (e) { ok = e === null; }
console.log(ok);
"#,
        ["true"]
    };

    throw_undefined_is_caught_as_undefined => {
        r#"
let ok = false;
try { throw undefined; } catch (e) { ok = e === undefined; }
console.log(ok);
"#,
        ["true"]
    };

    throw_boolean_is_caught_by_identity => {
        r#"
let caught;
try { throw true; } catch (e) { caught = e; }
console.log(caught === true);
"#,
        ["true"]
    };

    throw_symbol_is_caught_by_identity => {
        r#"
const s = Symbol("boom");
let same = false;
try { throw s; } catch (e) { same = e === s; }
console.log(same);
"#,
        ["true"]
    };

    throw_bigint_is_caught_by_value => {
        r#"
let caught;
try { throw 10n; } catch (e) { caught = e; }
console.log(caught === 10n);
console.log(typeof caught);
"#,
        ["true", "bigint"]
    };

    // ── 2. Error property descriptors (§20.5) ───────────────────────────

    error_message_is_not_enumerable => {
        // `message` is an own, non-enumerable data property — it must not
        // appear in Object.keys / for-in.
        r#"
const e = new Error("x");
console.log(Object.keys(e).length);
"#,
        ["0"]
    };

    error_instance_constructor_is_its_type => {
        r#"
console.log(new TypeError("x").constructor === TypeError);
console.log(new Error("x").constructor === Error);
"#,
        ["true", "true"]
    };

    error_name_lives_on_prototype_not_instance => {
        // `name` is inherited from the prototype, so it is not an own key.
        r#"
const e = new RangeError("x");
console.log(Object.prototype.hasOwnProperty.call(e, "name"));
console.log(e.name);
"#,
        ["false", "RangeError"]
    };

    // ── 3. finally runs on break / continue out of a loop try ───────────

    finally_runs_on_break_out_of_loop => {
        r#"
const log = [];
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) break;
        log.push("try" + i);
    } finally {
        log.push("fin" + i);
    }
}
console.log(log.join(","));
"#,
        ["try0,fin0,fin1"]
    };

    finally_runs_on_continue_in_loop => {
        r#"
const log = [];
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) continue;
        log.push("try" + i);
    } finally {
        log.push("fin" + i);
    }
}
console.log(log.join(","));
"#,
        ["try0,fin0,fin1,try2,fin2"]
    };

    throw_function_is_caught_and_callable => {
        r#"
let res = "";
try {
    throw function() { return "called"; };
} catch (e) {
    res = e();
}
console.log(res);
"#,
        ["called"]
    };
}
