// vybe-test: js/type_checking_patterns/null_undefined_check
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

function isNullOrUndefined(val) { return val == null; }
__check(__line(isNullOrUndefined(null)), "true");
__check(__line(isNullOrUndefined(undefined)), "true");
__check(__line(isNullOrUndefined(0)), "false");
__check(__line(isNullOrUndefined("")), "false");
__check(__line(isNullOrUndefined(false)), "false");
