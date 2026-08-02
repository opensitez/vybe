// vybe-test: js/number_precision_edge/number_max_min_value
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

__check(__line(Number.MAX_VALUE > 0), "true");
__check(__line(Number.MIN_VALUE > 0), "true");
__check(__line(Number.MIN_VALUE < 0.001), "true");
__check(__line(Number.POSITIVE_INFINITY === Infinity), "true");
__check(__line(Number.NEGATIVE_INFINITY === -Infinity), "true");
