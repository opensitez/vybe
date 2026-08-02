// vybe-test: js/number_precision_edge/floating_point_rounding_error
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

__check(__line(0.1 + 0.2 === 0.3), "false");
__check(__line(0.1 + 0.2), "0.30000000000000004");
