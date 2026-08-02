// vybe-test: js/function_constructor_dynamic_code_creation/test_js_async_function_constructor
// origin: languages/js/tests/js/test_js_function_constructor_dynamic_code_creation.rs

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
const asyncAdd = new AsyncFunction("a", "b", "return await Promise.resolve(a + b);");
(async () => {
    console.log(await asyncAdd(15, 25));
})();
