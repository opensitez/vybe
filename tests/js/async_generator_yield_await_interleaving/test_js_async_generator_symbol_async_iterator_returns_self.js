// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_symbol_async_iterator_returns_self
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

async function* gen() {}
const ag = gen();
__check(__line(ag[Symbol.asyncIterator]() === ag), "true");
