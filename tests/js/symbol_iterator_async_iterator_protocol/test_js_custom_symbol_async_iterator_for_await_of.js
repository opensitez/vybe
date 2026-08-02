// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_custom_symbol_async_iterator_for_await_of
// origin: languages/js/tests/js/test_js_symbol_iterator_async_iterator_protocol.rs

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

const asyncSeq = {
    [Symbol.asyncIterator]() {
        let i = 1;
        return {
            async next() {
                if (i <= 3) {
                    return { value: await Promise.resolve(i++ * 10), done: false };
                }
                return { done: true };
            }
        };
    }
};
(async () => {
    const res = [];
    for await (const val of asyncSeq) res.push(val);
    console.log(res.join(","));
})();
