// vybe-test: js/number_precision_edge/to_fixed_rounds_decimal
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

__check(__line((3.14159).toFixed(2)), "3.14");
__check(__line((1.005).toFixed(2)), "1.00"); // IEEE 754 — might be 1.00 or 1.01
__check(__line((100).toFixed(2)), "100.00");
