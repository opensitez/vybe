// vybe-test: js/class_extends_super_constructor_call/test_js_class_extends_builtin_error_subclassing
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class HttpError extends Error {
    constructor(status, message) {
        super(message);
        this.name = "HttpError";
        this.status = status;
    }
}
const err = new HttpError(404, "Not Found");
__check(__line(err.name + ":" + err.status + "|" + err.message + "|" + (err instanceof Error)), "HttpError:404|Not Found|true");
