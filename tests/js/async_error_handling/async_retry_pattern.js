// vybe-test: js/async_error_handling/async_retry_pattern
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

async function withRetry(fn, maxAttempts) {
    let lastError;
    for (let i = 0; i < maxAttempts; i++) {
        try { return await fn(i); }
        catch (e) { lastError = e; }
    }
    throw lastError;
}

let attempt = 0;
async function flakyOp(i) {
    attempt++;
    if (attempt < 3) throw new Error("fail");
    return "success";
}

withRetry(flakyOp, 5).then(r => {
    console.log(r);
    console.log(attempt);
});
