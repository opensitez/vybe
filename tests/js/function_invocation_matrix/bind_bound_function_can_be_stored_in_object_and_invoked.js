// vybe-test: js/function_invocation_matrix/bind_bound_function_can_be_stored_in_object_and_invoked
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

function mult(a, b) {
    return a * b;
}
const obj = { fn: mult.bind(null, 6) };
__check(__line(obj.fn(7)), "42");
