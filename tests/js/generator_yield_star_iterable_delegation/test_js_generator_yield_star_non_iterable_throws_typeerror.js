// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_non_iterable_throws_typeerror
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
    yield* 12345;
}
try {
    [...gen()];
} catch (e) {
    __check(__line("yield* Non-Iterable TypeError"), "yield* Non-Iterable TypeError");
}
