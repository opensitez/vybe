// vybe-test: js/function_invocation_matrix/bind_can_chain_bound_arguments
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

function parts(a, b, c) {
    return [a, b, c].join(":");
}
const one = parts.bind(null, "x");
const two = one.bind(null, "y");
__check(__line(two("z")), "x:y:z");
