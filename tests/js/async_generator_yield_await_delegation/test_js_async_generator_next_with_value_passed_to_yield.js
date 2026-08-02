// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_next_with_value_passed_to_yield
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

async function* echo() {
    const first = yield "Ready";
    yield `Echo: ${first}`;
}
(async () => {
    const gen = echo();
    await gen.next(); // Start generator
    const r2 = await gen.next("Hello");
    console.log(r2.value);
})();
