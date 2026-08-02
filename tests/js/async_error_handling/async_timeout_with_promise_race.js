// vybe-test: js/async_error_handling/async_timeout_with_promise_race
// origin: languages/js/tests/js/test_async_error_handling.rs

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

function timeout(ms, reason) {
    return new Promise((_, reject) =>
        setTimeout(() => reject(new Error(reason)), ms)
    );
}
async function withTimeout(fn, ms) {
    return Promise.race([fn(), timeout(ms, "timeout")]);
}

const fast = () => Promise.resolve("done");
withTimeout(fast, 1000).then(v => console.log(v));
