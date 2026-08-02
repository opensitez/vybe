// vybe-test: js/for_await_of_async_iterable_loop/for_await_of_prefers_symbol_async_iterator_over_sync
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

const dual = {
    [Symbol.iterator]() {
        return { next() { return { value: "sync", done: false }; } };
    },
    [Symbol.asyncIterator]() {
        return { async next() { return { value: "async", done: false }; } };
    }
};
(async () => {
    for await (const x of dual) {
        console.log(x);
        break;
    }
})();
