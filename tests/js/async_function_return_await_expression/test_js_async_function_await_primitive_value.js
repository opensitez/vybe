// vybe-test: js/async_function_return_await_expression/test_js_async_function_await_primitive_value
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

async function getPrimitive() {
    const val = await 100; // Primitive is wrapped in resolved promise
    return val * 2;
}
getPrimitive().then(res => console.log(res));
