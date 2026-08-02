// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_dynamic_instantiation
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

const AsyncGenFn = Object.getPrototypeOf(async function*(){}).constructor;
const genFn = new AsyncGenFn("yield await Promise.resolve('DynamicAsyncGen');");
(async () => {
    const gen = genFn();
    console.log((await gen.next()).value);
})();
