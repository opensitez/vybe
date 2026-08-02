// vybe-test: js/async_utility_patterns/async_filter
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

async function asyncFilter(arr, pred) {
    const results = await Promise.all(arr.map(async (item) => ({
        item,
        keep: await pred(item)
    })));
    return results.filter(r => r.keep).map(r => r.item);
}
async function main() {
    const evens = await asyncFilter([1, 2, 3, 4, 5, 6], async x => x % 2 === 0);
    console.log(evens.join(","));
}
main();
