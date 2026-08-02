// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_async_iterator_property_not_function_throws
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

const invalid = {
    [Symbol.asyncIterator]: 123
};
(async () => {
    try {
        for await (const _ of invalid) {}
    } catch (e) {
        console.log("Symbol.asyncIterator Not Callable TypeError");
    }
})();
