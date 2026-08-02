// vybe-test: js/generator_function_prototype/generator_function_prototype_call_passes_arguments
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

function* pick(a, b) { yield b; } __check(__line(GeneratorFunction.prototype.call.call(pick, null, 1, 9).next().value), "9");
