// vybe-test: js/function_invocation_matrix/function_prototype_call_can_invoke_borrowed_method_directly
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

const slice = String.prototype.slice;
__check(__line(Function.prototype.call.call(slice, "hello", 1, 4)), "ell");
