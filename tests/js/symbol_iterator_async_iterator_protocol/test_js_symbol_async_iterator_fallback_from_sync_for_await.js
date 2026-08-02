// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_async_iterator_fallback_from_sync_for_await
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

const syncIterable = [100, 200];
(async () => {
    const res = [];
    for await (const val of syncIterable) { // for-await-of falls back to Symbol.iterator wrapped in Promise if asyncIterator is missing
        res.push(val);
    }
    console.log(res.join(","));
})();
