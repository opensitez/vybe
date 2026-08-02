// vybe-test: js/async_await_deep/async_parallel_execution
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

async function delay(ms, val) {
    return new Promise(resolve => setTimeout(() => resolve(val), ms));
}
async function main() {
    const start = Date.now();
    const [a, b] = await Promise.all([
        delay(10, 1),
        delay(10, 2),
    ]);
    console.log(a + b);
}
main();
