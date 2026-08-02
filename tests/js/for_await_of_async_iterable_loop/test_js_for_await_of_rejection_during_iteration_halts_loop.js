// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_rejection_during_iteration_halts_loop
// origin: languages/js/tests/js/test_js_for_await_of_async_iterable_loop.rs

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

(async () => {
    const items = [Promise.resolve(1), Promise.reject("IterFail"), Promise.resolve(3)];
    const log = [];
    try {
        for await (const item of items) {
            log.push(item);
        }
    } catch (e) {
        log.push("Caught:" + e);
    }
    console.log(log.join(","));
})();
