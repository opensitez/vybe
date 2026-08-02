// vybe-test: js/ecma_functions/recursive_named_function_expression_internal_name
// origin: languages/js/tests/js/test_ecma_functions.rs

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

let outer = function inner(n) {
    if (n <= 1) return 1;
    return n * inner(n - 1);
};
__check(__line(outer(4)), "24");
