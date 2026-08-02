// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_result_expression_in_operator
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

function* inner() { return 100; }
function* outer() {
    const val = (yield* inner()) * 2;
    yield val;
}
__check(__line([...outer()][0]), "200");
