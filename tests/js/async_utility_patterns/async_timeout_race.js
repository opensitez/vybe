// vybe-test: js/async_utility_patterns/async_timeout_race
// origin: languages/js/tests/js/test_async_utility_patterns.rs

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

function timeout(ms) {
    return new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Timeout")), ms)
    );
}
async function withTimeout(promise, ms) {
    return Promise.race([promise, timeout(ms)]);
}
async function main() {
    const fast = Promise.resolve("done");
    const result = await withTimeout(fast, 5000);
    console.log(result);
}
main();
