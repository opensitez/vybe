// vybe-test: js/generator_function_prototype/generator_class_instance_method_instanceof_generator_function
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

class Box { *values() { yield 1; } } const b = new Box(); __check(__line(b.values instanceof GeneratorFunction), "true");
