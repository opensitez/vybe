// vybe-test: js/function_deep/apply_passes_args_as_array
// origin: languages/js/tests/js/test_function_deep.rs

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

function sum(a, b, c) { return a + b + c; }
__check(__line(sum.apply(null, [1, 2, 3])), "6");
