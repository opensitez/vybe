// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_array_of_promises
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
    const promises = [Promise.resolve("A"), Promise.resolve("B")];
    const items = [];
    for await (const item of promises) {
        items.push(item);
    }
    console.log(items.join(","));
})();
