// vybe-test: js/function_invocation_matrix/rest_parameter_does_not_change_arguments_length
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

function f(a, ...rest) {
    __check(__line(arguments.length), "3");
    __check(__line(rest.length), "2");
}
f(1, 2, 3);
