// vybe-test: js/async_utility_patterns/async_batch_processing_pattern
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

async function batchProcess(arr, size, fn) {
    const res = [];
    for (let i = 0; i < arr.length; i += size) {
        const batch = arr.slice(i, i + size);
        const batchRes = await Promise.all(batch.map(fn));
        res.push(...batchRes);
    }
    return res;
}
async function main() {
    const out = await batchProcess([1, 2, 3, 4, 5], 2, async x => x * 10);
    console.log(out.join(","));
}
main();
