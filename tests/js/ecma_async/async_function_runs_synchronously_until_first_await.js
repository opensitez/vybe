// vybe-test: js/ecma_async/async_function_runs_synchronously_until_first_await
// origin: languages/js/tests/js/test_ecma_async.rs

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

async function demo() {
    console.log("start");
    await 1;
    console.log("end");
}
console.log("before");
demo();
console.log("after");
