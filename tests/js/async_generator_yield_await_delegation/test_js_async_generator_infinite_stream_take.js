// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_infinite_stream_take
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

async function* infiniteNumbers() {
    let i = 1;
    while (true) {
        yield i++;
    }
}
(async () => {
    const results = [];
    for await (const n of infiniteNumbers()) {
        results.push(n);
        if (n === 3) break;
    }
    console.log(results.join(","));
})();
