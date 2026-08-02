// vybe-test: js/array_from_async/from_async_with_map_function
// origin: languages/js/tests/js/test_array_from_async.rs

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

async function* nums() { yield 1; yield 2; yield 3; }
Array.fromAsync(nums(), x => x * x).then(arr => console.log(arr.join(",")));
