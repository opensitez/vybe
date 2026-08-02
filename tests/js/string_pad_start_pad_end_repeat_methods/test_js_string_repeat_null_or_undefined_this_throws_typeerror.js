// vybe-test: js/string_pad_start_pad_end_repeat_methods/test_js_string_repeat_null_or_undefined_this_throws_typeerror
// origin: languages/js/tests/js/test_js_string_pad_start_pad_end_repeat_methods.rs

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
    String.prototype.repeat.call(undefined, 3);
} catch (e) {
    __check(__line("repeat Undefined This TypeError"), "repeat Undefined This TypeError");
}
