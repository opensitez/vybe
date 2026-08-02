// vybe-test: js/array_from_async/from_async_with_array_of_promises
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

const promises = [
    Promise.resolve(10),
    Promise.resolve(20),
    Promise.resolve(30),
];
Array.fromAsync(promises).then(arr => console.log(arr.join(",")));
