// vybe-test: js/number_precision_edge/number_to_string_removes_trailing_zeros
// origin: languages/js/tests/js/test_number_precision_edge.rs

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

__check(__line((1.0).toString()), "1");
__check(__line((1.50).toString()), "1.5");
__check(__line((1.500000).toString()), "1.5");
