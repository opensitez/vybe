// vybe-test: js/structured_clone_error_types_serialization/test_js_structured_clone_error_with_non_serializable_cause_throws
// origin: languages/js/tests/js/test_js_structured_clone_error_types_serialization.rs

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

const err = new Error("BadCauseMsg", { cause: () => {} });
try {
    structuredClone(err);
} catch (e) {
    __check(__line("DataCloneError Function Cause"), "DataCloneError Function Cause");
}
