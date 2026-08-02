// vybe-test: js/promise_resolve_reject_deferred_execution/test_js_promise_resolve_existing_promise_returns_same_instance
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

const p = Promise.resolve("Original");
console.log(Promise.resolve(p) === p);
