// vybe-test: js/async_generator_function_prototype/async_generator_function_bound_instanceof_after_partial_bind
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

async function* f(a, b) { yield a + b; } const partial = f.bind(null, 2); __check(__line(partial instanceof AsyncGeneratorFunction), "true");
