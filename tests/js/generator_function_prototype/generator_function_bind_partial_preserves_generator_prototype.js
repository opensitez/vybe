// vybe-test: js/generator_function_prototype/generator_function_bind_partial_preserves_generator_prototype
// origin: languages/js/tests/js/test_generator_function_prototype.rs

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

function* add(a, b) { yield a + b; } const plusOne = add.bind(null, 1); __check(__line(Object.getPrototypeOf(plusOne) === GeneratorFunction.prototype), "true");
