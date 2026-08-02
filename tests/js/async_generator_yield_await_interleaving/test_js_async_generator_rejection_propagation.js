// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_rejection_propagation
// origin: languages/js/tests/js/test_js_async_generator_yield_await_interleaving.rs

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

async function* gen() {
    yield 1;
    throw new Error("AsyncGenFailed");
}
(async () => {
    const ag = gen();
    await ag.next();
    try {
        await ag.next();
    } catch (e) {
        console.log("Caught: " + e.message);
    }
})();
