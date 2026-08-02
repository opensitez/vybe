// vybe-test: js/async_utility_patterns/promise_retry_with_limit
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

async function retry(fn, times) {
    for (let i = 0; i < times; i++) {
        try { return await fn(); }
        catch (e) { if (i === times - 1) throw e; }
    }
}
let attempts = 0;
async function main() {
    const result = await retry(async () => {
        attempts++;
        if (attempts < 3) throw new Error("not yet");
        return "success after " + attempts;
    }, 5);
    console.log(result);
    console.log(attempts);
}
main();
