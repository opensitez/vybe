// vybe-test: js/promise_unhandled_rejection_tracking/test_js_promise_finally_does_not_swallow_rejection
// origin: languages/js/tests/js/test_js_promise_unhandled_rejection_tracking.rs

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

Promise.reject("CriticalFailure")
    .finally(() => console.log("Logging Failure"))
    .catch(err => console.log("Caught After Finally: " + err));
