// vybe-test: js/async_await_deep/async_sequential_accumulator
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

async function accumulate(fns) {
    let result = 0;
    for (const fn of fns) {
        result = await fn(result);
    }
    return result;
}
async function main() {
    const result = await accumulate([
        async x => x + 1,
        async x => x * 2,
        async x => x + 10,
    ]);
    console.log(result); // ((0+1)*2)+10 = 12
}
main();
