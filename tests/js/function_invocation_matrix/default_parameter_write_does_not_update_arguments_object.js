// vybe-test: js/function_invocation_matrix/default_parameter_write_does_not_update_arguments_object
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

function f(a = 1) {
    a = 7;
    __check(__line(arguments[0]), "5");
}
f(5);
