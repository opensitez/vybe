// vybe-test: js/function_constructor_dynamic_code_creation/test_js_async_generator_function_constructor
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

const AsyncGeneratorFunction = Object.getPrototypeOf(async function*(){}).constructor;
const asyncGen = new AsyncGeneratorFunction("a", "yield await Promise.resolve(a * 2);");
(async () => {
    const ag = asyncGen(10);
    console.log((await ag.next()).value);
})();
