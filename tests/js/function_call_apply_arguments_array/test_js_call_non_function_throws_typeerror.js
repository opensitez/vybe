// vybe-test: js/function_call_apply_arguments_array/test_js_call_non_function_throws_typeerror
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

try {
    Function.prototype.call.call("not_a_function");
} catch (e) {
    __check(__line("call Non-Function TypeError"), "call Non-Function TypeError");
}
