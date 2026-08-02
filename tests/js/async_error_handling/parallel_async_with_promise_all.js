// vybe-test: js/async_error_handling/parallel_async_with_promise_all
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

async function delay(n) {
    await Promise.resolve();
    return n * n;
}
async function main() {
    const results = await Promise.all([delay(2), delay(3), delay(4)]);
    console.log(results.join(","));
}
main();
