// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_yield_star_async_iterable
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

async function* inner() {
    yield await Promise.resolve("A");
    yield await Promise.resolve("B");
}
async function* outer() {
    yield* inner();
}
(async () => {
    const items = [];
    for await (const x of outer()) items.push(x);
    console.log(items.join(","));
})();
