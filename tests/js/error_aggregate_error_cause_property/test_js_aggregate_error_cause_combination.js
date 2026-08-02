// vybe-test: js/error_aggregate_error_cause_property/test_js_aggregate_error_cause_combination
// origin: languages/js/tests/js/test_js_error_aggregate_error_cause_property.rs

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

const aggErr = new AggregateError([1, 2], "Operation failed", { cause: "RootCause" });
__check(__line(aggErr.message + "|cause=" + aggErr.cause + "|errors=" + aggErr.errors.join(",")), "Operation failed|cause=RootCause|errors=1,2");
