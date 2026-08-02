// vybe-test: js/function_length_name_properties_descriptors/test_js_function_caller_and_arguments_properties_in_strict_mode_throw
// origin: languages/js/tests/js/test_js_function_length_name_properties_descriptors.rs

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

function strictFn() {
    "use strict";
    try {
        strictFn.caller;
    } catch (e) {
        __check(__line("Strict Caller TypeError"), "Strict Caller TypeError");
    }
}
strictFn();
