// vybe-test: js/number_precision_edge/to_precision_significant_digits
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

__check(__line((123.456).toPrecision(5)), "123.46");
__check(__line((0.000123).toPrecision(2)), "0.00012");
__check(__line((1).toPrecision(4)), "1.000");
