// vybe-test: js/async_await_deep/async_function_returns_rejected_promise_handled_in_catch
// origin: languages/js/tests/js/test_async_await_deep.rs

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

async function fail() {
    return Promise.reject(new Error("rejected_return"));
}
async function main() {
    try {
        await fail();
    } catch (e) {
        console.log(e.message);
    }
}
main();
