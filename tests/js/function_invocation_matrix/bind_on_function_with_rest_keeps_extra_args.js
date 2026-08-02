// vybe-test: js/function_invocation_matrix/bind_on_function_with_rest_keeps_extra_args
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

function pack(head, ...rest) {
    return head + ":" + rest.join(",");
}
const fn = pack.bind(null, "a");
__check(__line(fn("b", "c")), "a:b,c");
