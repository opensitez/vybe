// vybe-test: js/promise_patterns_deep/async_sequential_order
// origin: languages/js/tests/js/test_promise_patterns_deep.rs

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
async function step(n) {
    log.push("start " + n);
    await Promise.resolve();
    log.push("end " + n);
    return n;
}
async function main() {
    await step(1);
    await step(2);
    console.log(log.join(","));
}
main();
