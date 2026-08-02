// vybe-test: js/async_generator_function_prototype/bound_async_generator_function_prototype_is_async_generator_function_prototype
// origin: languages/js/tests/js/test_async_generator_function_prototype.rs

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

async function* value() { yield 1; } const b = value.bind(null); __check(__line(Object.getPrototypeOf(b) === AsyncGeneratorFunction.prototype), "true");
