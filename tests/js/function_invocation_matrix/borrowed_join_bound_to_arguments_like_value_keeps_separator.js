// vybe-test: js/function_invocation_matrix/borrowed_join_bound_to_arguments_like_value_keeps_separator
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

function f() {
    const join = Array.prototype.join.bind(arguments, "/");
    __check(__line(join()), "x/y/z");
}
f("x", "y", "z");
