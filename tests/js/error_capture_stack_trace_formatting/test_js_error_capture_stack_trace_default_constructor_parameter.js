// vybe-test: js/error_capture_stack_trace_formatting/test_js_error_capture_stack_trace_default_constructor_parameter
// origin: languages/js/tests/js/test_js_error_capture_stack_trace_formatting.rs

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

class BaseErr extends Error {
    constructor(msg) {
        super(msg);
        Error.captureStackTrace(this, this.constructor);
    }
}
const e = new BaseErr("BaseMsg");
__check(__line(e.stack.includes("BaseErr: BaseMsg")), "true");
