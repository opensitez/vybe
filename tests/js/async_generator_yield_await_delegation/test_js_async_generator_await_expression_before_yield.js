// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_await_expression_before_yield
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

async function* asyncData() {
    const a = await Promise.resolve(10);
    yield a * 2;
    const b = await Promise.resolve(20);
    yield b * 2;
}
(async () => {
    const gen = asyncData();
    const v1 = (await gen.next()).value;
    const v2 = (await gen.next()).value;
    console.log(`${v1},${v2}`);
})();
