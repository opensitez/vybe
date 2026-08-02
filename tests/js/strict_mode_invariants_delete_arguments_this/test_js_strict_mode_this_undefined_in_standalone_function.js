// vybe-test: js/strict_mode_invariants_delete_arguments_this/test_js_strict_mode_this_undefined_in_standalone_function
// origin: languages/js/tests/js/test_js_strict_mode_invariants_delete_arguments_this.rs

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

function getThis() {
    "use strict";
    return this;
}
__check(__line(getThis() === undefined), "true");
