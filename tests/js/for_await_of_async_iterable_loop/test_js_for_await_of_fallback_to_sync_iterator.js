// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_fallback_to_sync_iterator
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

const syncIterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 2 ? { value: ++i, done: false } : { done: true };
            }
        };
    }
};
(async () => {
    const res = [];
    for await (const x of syncIterable) {
        res.push(x);
    }
    console.log(res.join(","));
})();
