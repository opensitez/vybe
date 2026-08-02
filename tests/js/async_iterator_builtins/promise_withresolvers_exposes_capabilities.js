// vybe-test: js/async_iterator_builtins/promise_withresolvers_exposes_capabilities
// origin: languages/js/tests/js/test_async_iterator_builtins.rs

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

const result = Promise.withResolvers();
console.log(typeof result.promise);
console.log(typeof result.resolve);
console.log(typeof result.reject);
