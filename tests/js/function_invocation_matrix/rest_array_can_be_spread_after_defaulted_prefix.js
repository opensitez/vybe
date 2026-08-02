// vybe-test: js/function_invocation_matrix/rest_array_can_be_spread_after_defaulted_prefix
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

function join(head = "x", ...tail) {
    __check(__line([head].concat(tail).join(",")), "x,a,b");
}
join(undefined, "a", "b");
