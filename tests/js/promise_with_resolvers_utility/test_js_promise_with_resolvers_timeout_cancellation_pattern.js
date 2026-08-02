// vybe-test: js/promise_with_resolvers_utility/test_js_promise_with_resolvers_timeout_cancellation_pattern
// origin: languages/js/tests/js/test_js_promise_with_resolvers_utility.rs

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

const { promise, resolve, reject } = Promise.withResolvers();
const timer = setTimeout(() => reject("TimeoutError"), 1000);

resolve("DataArrivedFast");
clearTimeout(timer);

(async () => {
    console.log(await promise);
})();
