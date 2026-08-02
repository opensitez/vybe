// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_empty_generator_returns_value
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

function* emptyGen() { return "EmptyVal"; }
function* outer() {
    const res = yield* emptyGen();
    yield res;
}
__check(__line([...outer()].join(",")), "EmptyVal");
