// vybe-test: js/async_utility_patterns/async_reduce_sequential
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

async function asyncReduce(arr, fn, init) {
    let acc = init;
    for (const item of arr) {
        acc = await fn(acc, item);
    }
    return acc;
}
async function main() {
    const result = await asyncReduce([1, 2, 3, 4, 5], async (acc, x) => acc + x, 0);
    console.log(result);
}
main();
