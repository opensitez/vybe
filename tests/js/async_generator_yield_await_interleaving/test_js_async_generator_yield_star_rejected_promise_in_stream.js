// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_yield_star_rejected_promise_in_stream
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
    yield 1;
    throw new Error("StreamError");
}
async function* outer() {
    yield* inner();
}
(async () => {
    try {
        for await (const _ of outer());
    } catch (e) {
        console.log(e.message);
    }
})();
