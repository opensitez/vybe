// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_passing_arguments_into_next
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

function* inner() {
    const x = yield 1;
    const y = yield x * 2;
    return y * 3;
}
function* outer() {
    const ret = yield* inner();
    yield "outerRet:" + ret;
}
const g = outer();
__check(__line(g.next().value), "1"); // 1
__check(__line(g.next(10).value), "20"); // 20
__check(__line(g.next(5).value), "outerRet:15"); // "outerRet:15"
