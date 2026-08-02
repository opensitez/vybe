// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_throw_method_injection
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

async function* errorHandler() {
    try {
        yield "Initial";
    } catch (e) {
        yield "Handled: " + e;
    }
}
(async () => {
    const gen = errorHandler();
    await gen.next();
    const r2 = await gen.throw("InjectedError");
    console.log(r2.value);
})();
