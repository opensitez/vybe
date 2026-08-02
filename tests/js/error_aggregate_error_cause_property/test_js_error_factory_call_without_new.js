// vybe-test: js/error_aggregate_error_cause_property/test_js_error_factory_call_without_new
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

const err = Error("NoNewKeyword");
__check(__line(err.message + "|" + (err instanceof Error)), "NoNewKeyword|true");
