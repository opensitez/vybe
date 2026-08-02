// vybe-test: js/async_function_return_await_expression/test_js_async_function_sequential_vs_parallel_await
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

async function testParallel() {
    const p1 = Promise.resolve("P1");
    const p2 = Promise.resolve("P2");
    const r1 = await p1;
    const r2 = await p2;
    return `${r1}+${r2}`;
}
testParallel().then(res => console.log(res));
