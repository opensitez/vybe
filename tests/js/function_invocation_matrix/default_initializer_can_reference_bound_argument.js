// vybe-test: js/function_invocation_matrix/default_initializer_can_reference_bound_argument
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

function f(a, b = a * 2) {
    return b;
}
const g = f.bind(null, 4);
__check(__line(g()), "8");
