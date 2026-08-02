// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_yield_star_delegated_throw_caught_in_outer
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

async function* inner() {
    yield 1;
    throw new Error("inner_err");
}
async function* outer() {
    try {
        yield* inner();
    } catch (e) {
        yield "caught:" + e.message;
    }
}
(async () => {
    const items = [];
    for await (const x of outer()) items.push(x);
    console.log(items.join(","));
})();
