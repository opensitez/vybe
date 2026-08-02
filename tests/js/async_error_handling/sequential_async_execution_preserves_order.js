// vybe-test: js/async_error_handling/sequential_async_execution_preserves_order
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

const log = [];
async function task(n) {
    log.push("start:" + n);
    await Promise.resolve();
    log.push("end:" + n);
    return n;
}
async function main() {
    await task(1);
    await task(2);
}
main().then(() => console.log(log.join(",")));
