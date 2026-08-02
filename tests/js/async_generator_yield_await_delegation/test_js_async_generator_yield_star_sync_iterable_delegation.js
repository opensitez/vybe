// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_yield_star_sync_iterable_delegation
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

async function* delegateSync() {
    yield* [10, 20, 30];
}
(async () => {
    const results = [];
    for await (const n of delegateSync()) {
        results.push(n * 2);
    }
    console.log(results.join(","));
})();
