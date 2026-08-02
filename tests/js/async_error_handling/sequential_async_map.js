// vybe-test: js/async_error_handling/sequential_async_map
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

async function asyncMap(arr, fn) {
    const results = [];
    for (const item of arr) {
        results.push(await fn(item));
    }
    return results;
}
asyncMap([1, 2, 3], async x => x * x)
    .then(r => console.log(r.join(",")));
