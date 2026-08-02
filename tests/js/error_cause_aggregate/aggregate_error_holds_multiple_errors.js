// vybe-test: js/error_cause_aggregate/aggregate_error_holds_multiple_errors
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

const err = new AggregateError([
    new Error("first"),
    new Error("second"),
    new Error("third")
], "Multiple errors occurred");
__check(__line(err.message), "Multiple errors occurred");
__check(__line(err.errors.length), "3");
__check(__line(err.errors[0].message), "first");
__check(__line(err.errors[2].message), "third");
