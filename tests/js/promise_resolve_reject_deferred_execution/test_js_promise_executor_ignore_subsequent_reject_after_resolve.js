// vybe-test: js/promise_resolve_reject_deferred_execution/test_js_promise_executor_ignore_subsequent_reject_after_resolve
// origin: languages/js/tests/js/test_js_promise_resolve_reject_deferred_execution.rs

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

new Promise((resolve, reject) => {
    resolve("Win");
    throw new Error("Ignored Exception After Resolve");
})
.then(res => console.log(res));
