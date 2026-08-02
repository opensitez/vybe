// vybe-test: js/number_bigint/bigint_loose_vs_strict_comparison_with_number
// origin: languages/js/tests/js/test_number_bigint.rs

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

// == coerces: 1n == 1 is true
__check(__line(1n == 1), "true");
// === does not coerce: 1n === 1 is false (different types)
__check(__line(1n === 1), "false");
