// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_concurrent_next_queue_ordering
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

async function* sequence() {
    yield 1;
    yield 2;
}
const gen = sequence();
// Calling next concurrently queues promises in FIFO order!
const p1 = gen.next();
const p2 = gen.next();
Promise.all([p1, p2]).then(results => {
    console.log(`${results[0].value},${results[1].value}`);
});
