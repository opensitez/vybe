// vybe-test: js/promise_microtasks/promise_executor_runs_synchronously
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
new Promise((resolve) => {
    log.push("executor");
    resolve();
});
log.push("after");
// microtasks haven't run yet, but executor already ran
console.log(log.join(","));
