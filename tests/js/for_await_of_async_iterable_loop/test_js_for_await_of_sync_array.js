// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_sync_array
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
    const results = [];
    for await (const x of [10, 20, 30]) {
        results.push(x * 2);
    }
    console.log(results.join(","));
})();
