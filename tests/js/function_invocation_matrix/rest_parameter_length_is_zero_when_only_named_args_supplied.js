// vybe-test: js/function_invocation_matrix/rest_parameter_length_is_zero_when_only_named_args_supplied
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

function f(a, b, ...rest) {
    __check(__line(rest.length), "0");
}
f(1, 2);
