// vybe-test: js/function_invocation_matrix/arguments_object_and_rest_collect_same_tail_values
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
    __check(__line(arguments[2]), "z");
    __check(__line(rest[1]), "z");
}
f("x", "y", "z");
