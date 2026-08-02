// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_destructured_bindings
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
    const pairs = [Promise.resolve([1, "a"]), Promise.resolve([2, "b"])];
    const log = [];
    for await (const [id, label] of pairs) {
        log.push(`${id}:${label}`);
    }
    console.log(log.join(","));
})();
