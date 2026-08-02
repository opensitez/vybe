// vybe-test: js/error_capture_stack_trace_formatting/test_js_error_prepare_stack_trace_v8_hook
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

const origPrepare = Error.prepareStackTrace;
Error.prepareStackTrace = (err, structuredStackTrace) => {
    return `CallSiteCount:${structuredStackTrace.length}`;
};
try {
    const err = new Error("HooksTest");
    __check(__line(err.stack.startsWith("CallSiteCount:")), "true");
} finally {
    Error.prepareStackTrace = origPrepare;
}
