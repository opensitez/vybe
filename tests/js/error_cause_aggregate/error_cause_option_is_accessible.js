// vybe-test: js/error_cause_aggregate/error_cause_option_is_accessible
// origin: languages/js/tests/js/test_error_cause_aggregate.rs

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

const original = new TypeError("original issue");
const wrapped = new Error("high-level failure", { cause: original });
__check(__line(wrapped.message), "high-level failure");
__check(__line(wrapped.cause instanceof TypeError), "true");
__check(__line(wrapped.cause.message), "original issue");
