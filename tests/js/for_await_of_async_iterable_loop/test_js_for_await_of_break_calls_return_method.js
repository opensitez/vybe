// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_break_calls_return_method
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

let returnCalled = false;
const customAsyncIterable = {
    [Symbol.asyncIterator]() {
        return {
            async next() { return { value: 1, done: false }; },
            async return() { returnCalled = true; return { done: true }; }
        };
    }
};
(async () => {
    for await (const x of customAsyncIterable) {
        break;
    }
    console.log("Return Method Called: " + returnCalled);
})();
