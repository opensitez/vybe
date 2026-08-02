// vybe-test: js/promise_microtasks/async_await_awaits_promise_chain
// origin: languages/js/tests/js/test_promise_microtasks.rs

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

async function compute() {
    const v1 = await Promise.resolve(10);
    const v2 = await Promise.resolve(v1 * 3);
    return v2;
}
compute().then(v => console.log(v));
