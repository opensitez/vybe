// vybe-test: js/strict_mode/strict_mode_undeclared_variable_throws
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

function f() {
    "use strict";
    let threw = false;
    try { x = 5; } catch (e) { threw = e instanceof ReferenceError; }
    return threw;
}
__check(__line(f()), "true");
