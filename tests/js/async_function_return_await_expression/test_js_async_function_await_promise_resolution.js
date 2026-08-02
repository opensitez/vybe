// vybe-test: js/async_function_return_await_expression/test_js_async_function_await_promise_resolution
// origin: languages/js/tests/js/test_js_async_function_return_await_expression.rs

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
    const a = await Promise.resolve(10);
    const b = await Promise.resolve(20);
    return a + b;
}
compute().then(res => console.log(res));
