// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_custom_async_iterable
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

const customAsyncIterable = {
    [Symbol.asyncIterator]() {
        let count = 0;
        return {
            async next() {
                if (count < 3) {
                    return { value: ++count, done: false };
                }
                return { value: undefined, done: true };
            }
        };
    }
};
(async () => {
    const vals = [];
    for await (const v of customAsyncIterable) {
        vals.push(v);
    }
    console.log(vals.join(","));
})();
