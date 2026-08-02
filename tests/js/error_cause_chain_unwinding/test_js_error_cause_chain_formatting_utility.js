// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_chain_formatting_utility
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

function formatErrorChain(err) {
    let msg = err.name + ": " + err.message;
    if (err.cause) {
        msg += "\n  [caused by]: " + (err.cause instanceof Error ? formatErrorChain(err.cause) : err.cause);
    }
    return msg;
}
const inner = new TypeError("Invalid argument");
const outer = new Error("Action failed", { cause: inner });
__check(__line(formatErrorChain(outer)), "Error: Action failed\n  [caused by]: TypeError: Invalid argument");
