// vybe-test: js/async_try_catch_finally_rejection_propagation/test_js_async_finally_throw_overrides_rejection
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

async function testThrowOverride() {
    try {
        await Promise.reject("OriginalError");
    } finally {
        throw new Error("FinallyError");
    }
}
testThrowOverride().catch(e => console.log(e.message));
