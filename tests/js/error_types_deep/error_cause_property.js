// vybe-test: js/error_types_deep/error_cause_property
// origin: languages/js/tests/js/test_error_types_deep.rs

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

const cause = new Error("root cause");
const err = new Error("outer", { cause });
__check(__line(err.message), "outer");
__check(__line(err.cause === cause), "true");
__check(__line(err.cause.message), "root cause");
