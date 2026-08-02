// vybe-test: js/async_generator_function_prototype/async_generator_function_prototype_apply_spreads_array_args
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

async function* pick(a, b) { yield a + b; } __check(__line(typeof AsyncGeneratorFunction.prototype.apply.call(pick, null, [4, 5]).next), "function");
