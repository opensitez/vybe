// vybe-test: js/bigint_ops/bigint_ordering_operators
// origin: languages/js/tests/js/test_bigint_ops.rs

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

__check(__line(10n < 20n), "true");
__check(__line(20n > 10n), "true");
__check(__line(10n <= 10n), "true");
__check(__line(10n >= 11n), "false");
