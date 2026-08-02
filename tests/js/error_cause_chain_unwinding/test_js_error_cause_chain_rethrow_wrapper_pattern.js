// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_chain_rethrow_wrapper_pattern
// origin: languages/js/tests/js/test_js_error_cause_chain_unwinding.rs

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

function runOperation() {
    try {
        JSON.parse("invalid_json");
    } catch (e) {
        throw new Error("Failed to parse config file", { cause: e });
    }
}
try {
    runOperation();
} catch (e) {
    __check(__line(e.message + "|isSyntaxError=" + (e.cause instanceof SyntaxError)), "Failed to parse config file|isSyntaxError=true");
}
