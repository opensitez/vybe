// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_return_method_early_completion
// origin: languages/js/tests/js/test_js_async_generator_yield_await_delegation.rs

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

async function* counter() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        console.log("Async Generator Cleanup");
    }
}
(async () => {
    const gen = counter();
    console.log((await gen.next()).value);
    const r2 = await gen.return("EarlyReturn");
    console.log(`${r2.value}|done=${r2.done}`);
})();
