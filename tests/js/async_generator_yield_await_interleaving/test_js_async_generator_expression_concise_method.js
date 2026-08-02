// vybe-test: js/async_generator_yield_await_interleaving/test_js_async_generator_expression_concise_method
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

const obj = {
    async *stream() {
        yield await Promise.resolve("S1");
        yield await Promise.resolve("S2");
    }
};
(async () => {
    const res = [];
    for await (const item of obj.stream()) res.push(item);
    console.log(res.join(","));
})();
