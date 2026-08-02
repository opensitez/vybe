// vybe-test: js/number_precision_edge/number_is_finite_vs_global
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

// Number.isFinite does NOT coerce
__check(__line(Number.isFinite(42)), "true");
__check(__line(Number.isFinite(Infinity)), "false");
__check(__line(Number.isFinite("42")), "false"); // false — no coercion
__check(__line(isFinite("42")), "true");        // true — coerces
