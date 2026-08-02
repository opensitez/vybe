// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_delegation_to_set_values
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

const set = new Set(["a", "b"]);
function* gen() {
    yield* set.values();
}
__check(__line([...gen()].join(",")), "a,b");
