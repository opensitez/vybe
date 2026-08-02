// vybe-test: js/suppressed_error_explicit_resource_management/test_js_suppressed_error_primitive_error_and_suppressed
// origin: languages/js/tests/js/test_js_suppressed_error_explicit_resource_management.rs

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

const err = new SuppressedError(404, "CleanupFailed");
__check(__line(err.error + "|" + err.suppressed), "404|CleanupFailed");
