// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_delegation_to_map_entries
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

const map = new Map([["k1", "v1"], ["k2", "v2"]]);
function* gen() {
    yield* map.entries();
}
__check(__line([...gen()].map(pair => pair.join("=")).join(",")), "k1=v1,k2=v2");
