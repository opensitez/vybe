// vybe-test: js/promises/async_await_with_promise_all
// origin: languages/js/tests/js/test_promises.rs

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

async function main() {
    let [a, b, c] = await Promise.all([
        Promise.resolve(10),
        Promise.resolve(20),
        Promise.resolve(30)
    ]);
    console.log(a + b + c);
}
main();
