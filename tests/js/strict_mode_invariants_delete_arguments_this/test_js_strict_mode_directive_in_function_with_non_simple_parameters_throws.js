// vybe-test: js/strict_mode_invariants_delete_arguments_this/test_js_strict_mode_directive_in_function_with_non_simple_parameters_throws
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

try {
    eval("function fn(a = 1) { 'use strict'; }"); // Non-simple parameters (defaults/destructuring) cannot have 'use strict'!
} catch (e) {
    __check(__line("Strict Directive Non-Simple Parameter SyntaxError"), "Strict Directive Non-Simple Parameter SyntaxError");
}
