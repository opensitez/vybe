// vybe-test: js/function_invocation_matrix/call_can_forward_multiple_arguments
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

function sum(a, b, c) {
    return this.base + a + b + c;
}
const ctx = { base: 1 };
__check(__line(sum.call(ctx, 2, 3, 4)), "10");
