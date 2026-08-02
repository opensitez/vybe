// vybe-test: js/suppressed_error_explicit_resource_management/test_js_suppressed_error_class_inheritance
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

class ResourceError extends SuppressedError {}
const err = new ResourceError("a", "b", "CustomResourceError");
__check(__line(err.name + "|" + err.message + "|isSuppressed=" + (err instanceof SuppressedError)), "ResourceError|CustomResourceError|isSuppressed=true");
