// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_yield_star_async_iterable_delegation
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

async function* subGen() {
    yield "Sub1";
    yield "Sub2";
}
async function* mainGen() {
    yield "Start";
    yield* subGen();
    yield "End";
}
(async () => {
    const results = [];
    for await (const item of mainGen()) {
        results.push(item);
    }
    console.log(results.join(","));
})();
