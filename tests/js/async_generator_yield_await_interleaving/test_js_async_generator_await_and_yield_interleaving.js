// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_await_and_yield_interleaving
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

async function* asyncGen() {
    const a = await Promise.resolve(5);
    yield a * 2;
    const b = await Promise.resolve(10);
    yield a + b;
}
(async () => {
    const res = [];
    for await (const val of asyncGen()) res.push(val);
    console.log(res.join(","));
})();
