// vybe-test: js/async_patterns/promise_any_aggregate_error_errors_array
// origin: languages/js/tests/js/test_async_patterns.rs

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

Promise.any([Promise.reject("e1"), Promise.reject("e2")]).catch(err => {
    console.log(err.errors.join(","));
});
