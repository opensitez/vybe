// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_basic_yield_next
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

async function* generate() {
    yield 1;
    yield 2;
    yield 3;
}
(async () => {
    const gen = generate();
    const r1 = await gen.next();
    const r2 = await gen.next();
    const r3 = await gen.next();
    const r4 = await gen.next();
    console.log(`${r1.value},${r2.value},${r3.value}|done=${r4.done}`);
})();
