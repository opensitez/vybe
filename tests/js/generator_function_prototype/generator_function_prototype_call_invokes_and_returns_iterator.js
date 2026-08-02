// vybe-test: js/generator_function_prototype/generator_function_prototype_call_invokes_and_returns_iterator
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

function* f() { yield 1; } __check(__line(typeof GeneratorFunction.prototype.call.call(f, null).next), "function");
