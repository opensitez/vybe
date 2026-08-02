// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_try_finally_cleanup
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

let cleanedUp = false;
async function* gen() {
    try {
        yield 1;
    } finally {
        await Promise.resolve();
        cleanedUp = true;
    }
}
(async () => {
    const ag = gen();
    await ag.next();
    await ag.return();
    console.log(cleanedUp);
})();
