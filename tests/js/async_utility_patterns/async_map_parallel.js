// vybe-test: js/async_utility_patterns/async_map_parallel
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

async function mapParallel(arr, fn) {
    return Promise.all(arr.map(fn));
}
async function main() {
    const results = await mapParallel([1, 2, 3, 4], async x => x * x);
    console.log(results.join(","));
}
main();
