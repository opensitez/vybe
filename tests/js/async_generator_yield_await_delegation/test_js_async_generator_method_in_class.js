// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_method_in_class
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

class Streamer {
    async *items() {
        yield 100;
    }
}
(async () => {
    const gen = new Streamer().items();
    console.log((await gen.next()).value);
})();
