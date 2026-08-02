// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_method_in_object
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

const obj = {
    async *stream() {
        yield "A";
        yield "B";
    }
};
(async () => {
    const items = [];
    for await (const x of obj.stream()) items.push(x);
    console.log(items.join(""));
})();
