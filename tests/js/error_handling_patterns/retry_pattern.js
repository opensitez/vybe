// vybe-test: js/error_handling_patterns/retry_pattern
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            return await fn(attempt);
        } catch (e) {
            lastError = e;
        }
    }
    throw lastError;
}
let callCount = 0;
async function main() {
    const result = await withRetry(async (attempt) => {
        callCount++;
        if (attempt < 3) throw new Error("not yet");
        return "success";
    }, 3);
    console.log(result);
    console.log(callCount);
}
main();
