// vybe-test: js/async_await_deep/async_timeout_pattern
// origin: languages/js/tests/js/test_async_await_deep.rs

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

function withTimeout(promise, ms) {
    const timeout = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("timeout")), ms)
    );
    return Promise.race([promise, timeout]);
}
async function main() {
    // fast operation wins
    const fast = Promise.resolve("done");
    const result = await withTimeout(fast, 1000);
    console.log(result);
    // slow would timeout, but we test success path only
}
main();
