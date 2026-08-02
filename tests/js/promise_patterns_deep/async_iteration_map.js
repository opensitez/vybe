// vybe-test: js/promise_patterns_deep/async_iteration_map
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

async function asyncMap(arr, fn) {
    return Promise.all(arr.map(fn));
}
async function main() {
    const results = await asyncMap([1, 2, 3], async x => x * 2);
    console.log(results.join(","));
}
main();
