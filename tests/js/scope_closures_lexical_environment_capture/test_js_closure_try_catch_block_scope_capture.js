// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_try_catch_block_scope_capture
// origin: languages/js/tests/js/test_js_scope_closures_lexical_environment_capture.rs

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

let getError;
try {
    throw new Error("CaughtInTry");
} catch (err) {
    getError = () => err.message;
}
__check(__line(getError()), "CaughtInTry");
