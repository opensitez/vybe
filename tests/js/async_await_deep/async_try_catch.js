// vybe-test: js/async_await_deep/async_try_catch
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

async function unsafe() {
    throw new Error("oops");
}
async function main() {
    let caught = null;
    try {
        await unsafe();
    } catch (e) {
        caught = e.message;
    }
    console.log(caught);
}
main();
