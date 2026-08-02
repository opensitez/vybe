// vybe-test: js/function_invocation_matrix/default_parameter_uses_earlier_argument_only_when_tail_missing
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

function f(a, b = a + 1) {
    console.log(b);
}
f(2);
f(2, 10);
