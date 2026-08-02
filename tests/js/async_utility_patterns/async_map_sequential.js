// vybe-test: js/async_utility_patterns/async_map_sequential
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

async function mapSequential(arr, fn) {
    const results = [];
    for (const item of arr) {
        results.push(await fn(item));
    }
    return results;
}
async function main() {
    const results = await mapSequential([1, 2, 3], async x => x * 2);
    console.log(results.join(","));
}
main();
