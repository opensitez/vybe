// vybe-test: js/generator_function_prototype/generator_function_return_value_available_after_exhaustion
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

function* fin() { yield 1; return 9; } const it = fin(); it.next(); __check(__line(it.next().value), "9");
