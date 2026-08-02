// vybe-test: js/suppressed_error_explicit_resource_management/test_js_suppressed_error_constructor_properties
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

const primary = new Error("Primary Failure");
const suppressed = new Error("Cleanup Failure");
const err = new SuppressedError(primary, suppressed, "Resource Failure");

__check(__line(err.name + "|" + err.message + "|" + (err.error === primary) + "|" + (err.suppressed === suppressed)), "SuppressedError|Resource Failure|true|true");
