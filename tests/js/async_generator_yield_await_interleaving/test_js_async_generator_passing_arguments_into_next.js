// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_passing_arguments_into_next
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
    const x = yield "start";
    yield x * 10;
}
(async () => {
    const ag = gen();
    console.log((await ag.next()).value);
    console.log((await ag.next(5)).value);
})();
