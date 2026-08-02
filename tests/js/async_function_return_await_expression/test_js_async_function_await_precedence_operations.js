// vybe-test: js/async_function_return_await_expression/test_js_async_function_await_precedence_operations
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

async function testPrecedence() {
    const val = await Promise.resolve(10) + 5; // (await Promise.resolve(10)) + 5 = 15
    return val;
}
testPrecedence().then(res => console.log(res));
