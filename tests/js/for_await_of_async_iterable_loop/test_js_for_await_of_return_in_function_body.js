// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_return_in_function_body
// origin: languages/js/tests/js/test_js_for_await_of_async_iterable_loop.rs

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

async function findFirstEven(numbers) {
    for await (const n of numbers) {
        if (n % 2 === 0) return n;
    }
    return null;
}
findFirstEven([1, 3, 6, 7]).then(res => console.log(res));
