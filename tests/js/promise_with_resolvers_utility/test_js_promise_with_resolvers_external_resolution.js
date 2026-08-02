// vybe-test: js/promise_with_resolvers_utility/test_js_promise_with_resolvers_external_resolution
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

const { promise, resolve } = Promise.withResolvers();
resolve("ExternalValue");
(async () => {
    const val = await promise;
    console.log(val);
})();
