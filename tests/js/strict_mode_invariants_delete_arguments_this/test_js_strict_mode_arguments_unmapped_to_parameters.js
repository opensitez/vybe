// vybe-test: js/strict_mode_invariants_delete_arguments_this/test_js_strict_mode_arguments_unmapped_to_parameters
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

function fn(a) {
    "use strict";
    a = 100;
    return arguments[0]; // arguments[0] remains original passed value (not mapped to parameter 'a')!
}
__check(__line(fn(5)), "5");
