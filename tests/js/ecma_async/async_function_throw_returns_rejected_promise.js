// vybe-test: js/ecma_async/async_function_throw_returns_rejected_promise
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

async function err() {
    throw "boom";
}
err().catch(e => __check(__line(e), "boom"));
