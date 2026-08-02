// vybe-test: js/promise_unhandled_rejection_tracking/test_js_promise_late_handler_attachment
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

const p = Promise.reject("LateError");
// Attaching handler in subsequent microtask
Promise.resolve().then(() => {
    p.catch(err => console.log("Late Handled: " + err));
});
