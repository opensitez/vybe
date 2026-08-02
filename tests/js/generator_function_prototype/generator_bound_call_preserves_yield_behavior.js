// vybe-test: js/generator_function_prototype/generator_bound_call_preserves_yield_behavior
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

function* f(x) { yield x * 2; } const doubled = f.bind(null, 3); __check(__line(doubled().next().value), "6");
