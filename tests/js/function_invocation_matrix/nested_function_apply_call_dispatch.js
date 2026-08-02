// vybe-test: js/function_invocation_matrix/nested_function_apply_call_dispatch
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

function f(a, b) {
    return a + b;
}
__check(__line(Function.prototype.apply.call(f, null, [10, 20])), "30");
