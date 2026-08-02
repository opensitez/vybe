// vybe-test: js/function_invocation_matrix/arguments_reads_extra_values_beyond_named_parameters
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

function f(a) {
    __check(__line(arguments[1]), "y");
    __check(__line(arguments[2]), "z");
}
f("x", "y", "z");
