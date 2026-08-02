// vybe-test: js/try_catch_finally_return_override_control_flow/test_js_async_try_catch_finally_unwinding
// origin: languages/js/tests/js/test_js_try_catch_finally_return_override_control_flow.rs

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
    const log = [];
    try {
        log.push("AsyncTry");
        await Promise.reject(new Error("AsyncErr"));
    } catch (e) {
        log.push("AsyncCatch");
    } finally {
        log.push("AsyncFinally");
    }
    return log.join(",");
}
(async () => {
    console.log(await fn());
})();
