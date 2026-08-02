// vybe-test: js/strict_mode/strict_eval_has_own_variable_scope
// origin: languages/js/tests/js/test_strict_mode.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

"use strict";
eval("var evalVar = 42;");
let defined = false;
try { if (evalVar === 42) defined = true; } catch {}
// In strict mode, eval vars don't leak
console.log(!defined);
