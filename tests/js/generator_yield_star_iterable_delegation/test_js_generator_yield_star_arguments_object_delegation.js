// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_arguments_object_delegation
// origin: languages/js/tests/js/test_js_generator_yield_star_iterable_delegation.rs

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

function* gen() {
    yield* arguments;
}
function test(a, b) {
    return [...gen(a, b)];
}
__check(__line(test("first", "second").join(",")), "first,second");
