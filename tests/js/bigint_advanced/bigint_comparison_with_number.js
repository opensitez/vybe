// vybe-test: js/bigint_advanced/bigint_comparison_with_number
// origin: languages/js/tests/js/test_bigint_advanced.rs

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

__check(__line(1n < 2), "true");
__check(__line(2n > 1), "true");
__check(__line(1n == 1), "true");    // abstract equality
__check(__line(1n === 1), "false");   // strict: false (different types)
