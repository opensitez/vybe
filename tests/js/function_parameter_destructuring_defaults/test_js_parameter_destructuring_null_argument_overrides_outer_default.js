// vybe-test: js/function_parameter_destructuring_defaults/test_js_parameter_destructuring_null_argument_overrides_outer_default
// origin: languages/js/tests/js/test_js_function_parameter_destructuring_defaults.rs

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

function test(opts = { val: 10 }) {
    return opts;
}
__check(__line(test(null) === null), "true");
