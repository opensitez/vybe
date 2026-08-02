// vybe-test: js/async_patterns/async_generator_basic
// origin: languages/js/tests/js/test_async_patterns.rs

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

async function* asyncRange(start, end) {
    for (let i = start; i <= end; i++) {
        yield i;
    }
}
async function main() {
    for await (let n of asyncRange(1, 5)) {
        console.log(n);
    }
}
main();
