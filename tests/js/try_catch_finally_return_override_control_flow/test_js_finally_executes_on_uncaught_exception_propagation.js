// vybe-test: js/try_catch_finally_return_override_control_flow/test_js_finally_executes_on_uncaught_exception_propagation
// origin: languages/js/tests/js/test_js_try_catch_finally_return_override_control_flow.rs

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

let finallyExecuted = false;
function fn() {
    try {
        throw new Error("Uncaught");
    } finally {
        finallyExecuted = true;
    }
}
try {
    fn();
} catch (e) {
    __check(__line(e.message + "|FinallyExecuted=" + finallyExecuted), "Uncaught|FinallyExecuted=true");
}
