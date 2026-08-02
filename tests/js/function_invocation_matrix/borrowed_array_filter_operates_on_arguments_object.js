// vybe-test: js/function_invocation_matrix/borrowed_array_filter_operates_on_arguments_object
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

function f() {
    const out = Array.prototype.filter.call(arguments, x => x > 1);
    __check(__line(out.join(",")), "2,3");
}
f(1, 2, 3);
