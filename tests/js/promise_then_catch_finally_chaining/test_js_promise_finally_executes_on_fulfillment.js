// vybe-test: js/promise_then_catch_finally_chaining/test_js_promise_finally_executes_on_fulfillment
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

Promise.resolve("Success")
    .finally(() => console.log("Finally Done"))
    .then(res => console.log("Resolved: " + res));
