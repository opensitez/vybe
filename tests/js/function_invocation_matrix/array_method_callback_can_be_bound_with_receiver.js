// vybe-test: js/function_invocation_matrix/array_method_callback_can_be_bound_with_receiver
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

const ctx = { factor: 4 };
function mul(x) {
    return x * this.factor;
}
const out = [1, 2].map(mul.bind(ctx));
__check(__line(out.join(",")), "4,8");
