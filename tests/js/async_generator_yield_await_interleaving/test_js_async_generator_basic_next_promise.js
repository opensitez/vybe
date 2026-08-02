// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_basic_next_promise
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
    yield 10;
    yield 20;
}
(async () => {
    const ag = asyncGen();
    const p1 = await ag.next();
    const p2 = await ag.next();
    console.log(`${p1.value}:${p1.done} | ${p2.value}:${p2.done}`);
})();
