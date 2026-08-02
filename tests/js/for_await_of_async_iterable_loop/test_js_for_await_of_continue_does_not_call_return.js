// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_continue_does_not_call_return
// origin: languages/js/tests/js/test_js_for_await_of_async_iterable_loop.rs

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

(async () => {
    let returnCalls = 0;
    const iterable = {
        [Symbol.asyncIterator]() {
            let i = 0;
            return {
                async next() {
                    return i < 4 ? { value: ++i, done: false } : { done: true };
                },
                async return() {
                    returnCalls++;
                    return { done: true };
                }
            };
        }
    };

    const seen = [];
    for await (const n of iterable) {
        if (n % 2 === 0) continue;
        seen.push(n);
    }
    console.log(seen.join(","));
    console.log(returnCalls);
})();
