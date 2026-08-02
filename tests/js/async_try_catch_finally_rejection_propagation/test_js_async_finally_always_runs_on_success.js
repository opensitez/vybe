// vybe-test: js/async_try_catch_finally_rejection_propagation/test_js_async_finally_always_runs_on_success
// origin: languages/js/tests/js/test_js_async_try_catch_finally_rejection_propagation.rs

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

async function run() {
    try {
        return await Promise.resolve("Success");
    } finally {
        console.log("Cleanup Executed");
    }
}
run().then(res => console.log("Result: " + res));
