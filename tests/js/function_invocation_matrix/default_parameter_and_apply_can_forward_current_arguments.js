// vybe-test: js/function_invocation_matrix/default_parameter_and_apply_can_forward_current_arguments
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

function wrap(a = 1) {
    function inner(x, y) {
        __check(__line(x + y), "5");
    }
    inner.apply(null, arguments);
}
wrap(2, 3);
