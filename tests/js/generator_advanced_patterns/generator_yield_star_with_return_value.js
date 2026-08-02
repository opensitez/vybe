// vybe-test: js/generator_advanced_patterns/generator_yield_star_with_return_value
// origin: languages/js/tests/js/test_generator_advanced_patterns.rs

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

function* inner() {
    yield 1;
    yield 2;
    return "inner done";
}
function* outer() {
    const result = yield* inner();
    __check(__line(result), "inner done"); // return value of inner
    yield 3;
}
__check(__line([...outer()].join(",")), "1,2,3");
