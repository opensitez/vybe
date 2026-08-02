// vybe-test: js/promise_unhandled_rejection_tracking/test_js_promise_rejection_rethrown_in_catch
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

Promise.reject("Initial")
    .catch(err => {
        console.log("Stage 1: " + err);
        throw new Error("Rethrown");
    })
    .catch(err => console.log("Stage 2: " + err.message));
