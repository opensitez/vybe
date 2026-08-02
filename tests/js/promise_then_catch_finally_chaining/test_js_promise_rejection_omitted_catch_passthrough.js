// vybe-test: js/promise_then_catch_finally_chaining/test_js_promise_rejection_omitted_catch_passthrough
// origin: languages/js/tests/js/test_js_promise_then_catch_finally_chaining.rs

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

Promise.reject("Uncaught")
    .then(val => val) // Reject handler omitted
    .catch(err => console.log("PassThrough: " + err));
