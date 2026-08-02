// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_yield_star_sync_iterable
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
    yield* [1, 2, 3];
}
(async () => {
    const items = [];
    for await (const x of gen()) items.push(x);
    console.log(items.join(","));
})();
