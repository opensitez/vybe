// vybe-test: js/strict_mode/strict_mode_arguments_callee_throws_typeerror
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
function f() {
    try {
        return arguments.callee;
    } catch (e) {
        return e instanceof TypeError ? "typeerror" : "error";
    }
}
__check(__line(f()), "typeerror");
