// vybe-test: js/type_checking_patterns/number_type_checks
// origin: languages/js/tests/js/test_type_checking_patterns.rs

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

const checks = {
    int: n => Number.isInteger(n),
    finite: n => Number.isFinite(n),
    safe: n => Number.isSafeInteger(n),
    nan: n => Number.isNaN(n),
};
__check(__line(checks.int(5)), "true");
__check(__line(checks.int(5.5)), "false");
__check(__line(checks.finite(Infinity)), "false");
__check(__line(checks.safe(2**53)), "false");
__check(__line(checks.nan(NaN)), "true");
