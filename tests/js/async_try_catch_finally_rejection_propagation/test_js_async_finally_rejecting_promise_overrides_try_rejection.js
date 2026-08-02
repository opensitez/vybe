// vybe-test: js/async_try_catch_finally_rejection_propagation/test_js_async_finally_rejecting_promise_overrides_try_rejection
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

async function fn() {
    try {
        return Promise.reject("try_rej");
    } finally {
        return Promise.reject("fin_rej");
    }
}
fn().catch(e => console.log(e));
