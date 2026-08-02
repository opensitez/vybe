// vybe-test: js/number_edge_cases/tofixed_trailing_zeros
// origin: languages/js/tests/js/test_number_edge_cases.rs

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

__check(__line((1.0).toFixed(2)), "1.00");
__check(__line((1.5).toFixed(0)), "2");
__check(__line((0).toFixed(3)), "0.000");
