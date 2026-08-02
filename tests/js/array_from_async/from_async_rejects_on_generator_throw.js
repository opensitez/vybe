// vybe-test: js/array_from_async/from_async_rejects_on_generator_throw
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

async function* failing() {
    yield 1;
    throw new Error("boom");
}
Array.fromAsync(failing())
    .then(() => console.log("no"))
    .catch(e => console.log("caught:" + e.message));
