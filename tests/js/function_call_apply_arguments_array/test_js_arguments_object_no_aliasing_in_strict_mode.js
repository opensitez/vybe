// vybe-test: js/function_call_apply_arguments_array/test_js_arguments_object_no_aliasing_in_strict_mode
// origin: languages/js/tests/js/test_js_function_call_apply_arguments_array.rs

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

function mutate(a) {
    "use strict";
    a = 99; // Mutating parameter does NOT update arguments[0] in strict mode!
    return arguments[0];
}
__check(__line(mutate(10)), "10");
