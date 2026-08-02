// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_in_function_parameters
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

function fn(opts) {
    opts ||= {};
    opts.timeout ??= 1000;
    return opts.timeout;
}
__check(__line(fn() + "|" + fn({ timeout: 500 })), "1000|500");
