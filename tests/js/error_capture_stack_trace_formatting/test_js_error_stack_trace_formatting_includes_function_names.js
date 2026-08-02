// vybe-test: js/error_capture_stack_trace_formatting/test_js_error_stack_trace_formatting_includes_function_names
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

function levelA() { throw new Error("TraceErr"); }
function levelB() { levelA(); }

try {
    levelB();
} catch (e) {
    __check(__line(e.stack.includes("levelA") && e.stack.includes("levelB")), "true");
}
