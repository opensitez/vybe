// vybe-test: js/async_concurrency_patterns/timeout_with_abort
// origin: languages/js/tests/js/test_async_concurrency_patterns.rs

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
    let id;
    const timeout = new Promise((_, reject) => {
        id = setTimeout(() => reject(new Error("timeout")), ms);
    });
    return Promise.race([promise, timeout]).finally(() => clearTimeout(id));
}
async function main() {
    const fast = withTimeout(Promise.resolve("ok"), 1000);
    console.log(await fast);
    try {
        await withTimeout(new Promise(() => {}), 0);
    } catch(e) {
        console.log(e.message);
    }
}
main();
