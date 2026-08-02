// vybe-test: js/function_invocation_matrix/rest_collects_values_after_defaulted_head
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

function f(a = 1, ...rest) {
    __check(__line(a), "1");
    __check(__line(rest.join(",")), "2,3");
}
f(undefined, 2, 3);
