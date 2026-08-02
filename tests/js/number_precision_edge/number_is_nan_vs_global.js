// vybe-test: js/number_precision_edge/number_is_nan_vs_global
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

// Number.isNaN does NOT coerce
__check(__line(Number.isNaN(NaN)), "true");
__check(__line(Number.isNaN("NaN")), "false"); // false — no coercion
__check(__line(isNaN("NaN")), "true");        // true — coerces
__check(__line(Number.isNaN(undefined)), "false");
