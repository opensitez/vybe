// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_delegation_to_array
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
    yield* [10, 20, 30];
}
__check(__line([...gen()].join(",")), "10,20,30");
