// vybe-test: js/promise_microtasks/then_callbacks_run_after_current_task
// origin: languages/js/tests/js/test_promise_microtasks.rs

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

const log = [];
Promise.resolve().then(() => log.push("microtask"));
log.push("sync");
// After current sync code, microtask runs
Promise.resolve().then(() => console.log(log.join(",")));
