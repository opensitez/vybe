// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_queueing_next_calls
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
    yield 2;
}
(async () => {
    const ag = gen();
    const p1 = ag.next();
    const p2 = ag.next();
    const [r1, r2] = await Promise.all([p1, p2]);
    console.log(`${r1.value}:${r2.value}`);
})();
