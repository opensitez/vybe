// vybe-test: js/function_call_apply_bind/call_with_null_this_strict_fn
// origin: languages/js/tests/js/test_function_call_apply_bind.rs

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
function whatIsThis() { return this; }
__check(__line(whatIsThis.call(null)), "null");
__check(__line(whatIsThis.call(undefined)), "undefined");
