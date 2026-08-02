// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_iterator_not_a_function_throws_typeerror
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

const badIter = { [Symbol.iterator]: "not_a_function" };
try {
    for (const _ of badIter);
} catch (e) {
    console.log("Symbol.iterator Not Callable TypeError");
}
