// vybe-test: js/operators_deep/instanceof_rhs_coercion_with_non_function_is_false
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(3 instanceof Number), "false");
__check(__line(3 instanceof 123), "false");
__check(__line({} instanceof Number), "false");
__check(__line({} instanceof Object), "false");
