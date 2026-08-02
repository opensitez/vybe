// vybe-test: js/error_capture_stack_trace_formatting/test_js_error_stack_getter_does_not_throw_if_detached
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

const desc = Object.getOwnPropertyDescriptor(Error.prototype, "stack");
__check(__line(typeof desc === "undefined" || typeof desc.get === "function"), "true");
