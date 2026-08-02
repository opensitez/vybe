// vybe-test: js/promise_resolve_reject_deferred_execution/test_js_promise_resolve_self_cycle_rejection
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

const thenableCycle = {};
thenableCycle.then = (resolve) => resolve(thenableCycle);

Promise.resolve(thenableCycle).catch(err => console.log(err.name));
