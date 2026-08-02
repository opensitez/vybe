// vybe-test: js/async_function_return_await_expression/test_js_async_function_dynamic_instantiation
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

const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
const fn = new AsyncFunction("a", "b", "return (await a) + (await b);");
fn(Promise.resolve(5), Promise.resolve(10)).then(res => console.log(res));
