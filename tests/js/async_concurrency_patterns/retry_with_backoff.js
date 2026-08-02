// vybe-test: js/async_concurrency_patterns/retry_with_backoff
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

async function retry(fn, maxAttempts, delay = 0) {
    for (let i = 0; i < maxAttempts; i++) {
        try { return await fn(); }
        catch(e) {
            if (i === maxAttempts - 1) throw e;
            await new Promise(r => setTimeout(r, delay));
        }
    }
}
async function main() {
    let attempts = 0;
    const result = await retry(async () => {
        attempts++;
        if (attempts < 3) throw new Error("fail");
        return "success";
    }, 5);
    console.log(result);
    console.log(attempts);
}
main();
