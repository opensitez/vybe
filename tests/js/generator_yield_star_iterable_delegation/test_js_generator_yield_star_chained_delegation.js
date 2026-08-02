// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_chained_delegation
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

function* gen1() { yield 1; }
function* gen2() { yield* gen1(); yield 2; }
function* gen3() { yield* gen2(); yield 3; }
__check(__line([...gen3()].join(",")), "1,2,3");
