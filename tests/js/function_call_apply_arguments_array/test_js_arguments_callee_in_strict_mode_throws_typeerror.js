// vybe-test: js/function_call_apply_arguments_array/test_js_arguments_callee_in_strict_mode_throws_typeerror
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

function testCallee() {
    "use strict";
    try {
        arguments.callee;
    } catch (e) {
        __check(__line("Strict arguments.callee TypeError"), "Strict arguments.callee TypeError");
    }
}
testCallee();
