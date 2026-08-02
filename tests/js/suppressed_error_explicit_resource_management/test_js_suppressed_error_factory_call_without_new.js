// vybe-test: js/suppressed_error_explicit_resource_management/test_js_suppressed_error_factory_call_without_new
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

const err = SuppressedError("err1", "err2", "msg");
__check(__line(err.message + "|" + (err instanceof SuppressedError)), "msg|true");
