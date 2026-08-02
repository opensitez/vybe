// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_async_iterator_async_generator_concise_syntax
// origin: languages/js/tests/js/test_js_symbol_iterator_async_iterator_protocol.rs

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

class AsyncQueue {
    async *[Symbol.asyncIterator]() {
        yield await Promise.resolve("AQ1");
        yield await Promise.resolve("AQ2");
    }
}
(async () => {
    const items = [];
    for await (const x of new AsyncQueue()) items.push(x);
    console.log(items.join(","));
})();
