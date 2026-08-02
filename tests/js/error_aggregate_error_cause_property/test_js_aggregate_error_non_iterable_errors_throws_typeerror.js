// vybe-test: js/error_aggregate_error_cause_property/test_js_aggregate_error_non_iterable_errors_throws_typeerror
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

try {
    new AggregateError(12345, "Bad");
} catch (e) {
    __check(__line("AggregateError Non-Iterable TypeError"), "AggregateError Non-Iterable TypeError");
}
