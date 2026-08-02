// vybe-test: js/error_aggregate_error_cause_property/test_js_error_tostring_formatting
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

const err = new Error("SampleMessage");
__check(__line(err.toString()), "Error: SampleMessage");
