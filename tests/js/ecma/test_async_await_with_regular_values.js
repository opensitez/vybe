// vybe-test: js/ecma/test_async_await_with_regular_values
// origin: languages/js/tests/js/js_ecma_test.rs

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

async function getItems() {
            let a = await 10;
            let b = await 20;
            return a + b;
        }
        console.log(await getItems());
