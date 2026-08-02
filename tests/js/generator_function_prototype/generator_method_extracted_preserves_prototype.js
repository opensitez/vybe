// vybe-test: js/generator_function_prototype/generator_method_extracted_preserves_prototype
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

const src = { *emit() { yield 1; } }; const bare = src.emit; __check(__line(Object.getPrototypeOf(bare) === GeneratorFunction.prototype), "true");
