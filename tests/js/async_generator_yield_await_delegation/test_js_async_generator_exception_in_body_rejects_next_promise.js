// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_exception_in_body_rejects_next_promise
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

async function* thrower() {
    yield 1;
    throw new Error("GeneratorBodyError");
}
(async () => {
    const gen = thrower();
    await gen.next();
    try {
        await gen.next();
    } catch (e) {
        console.log(e.message);
    }
})();
