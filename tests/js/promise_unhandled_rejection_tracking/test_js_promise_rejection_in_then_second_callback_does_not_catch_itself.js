// vybe-test: js/promise_unhandled_rejection_tracking/test_js_promise_rejection_in_then_second_callback_does_not_catch_itself
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

Promise.reject("Err1")
    .then(
        val => val,
        err => { throw new Error("Err2"); }
    )
    .catch(err => console.log("CaughtInNextCatch: " + err.message));
