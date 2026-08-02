// vybe-test: js/error_aggregate_error_cause_property/test_js_aggregate_error_construction_and_errors_array
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

const err = new AggregateError([new Error("Err1"), new Error("Err2")], "Bulk Failure");
__check(__line(err.name + "|" + err.message + "|" + err.errors.length), "AggregateError|Bulk Failure|2");
