// vybe-test: js/function_bind_currying_bound_this/test_js_bound_generator_function_returns_generator
// origin: languages/js/tests/js/test_js_function_bind_currying_bound_this.rs

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

function* gen(a, b) {
    yield a * this.factor;
    yield b * this.factor;
}
const boundGen = gen.bind({ factor: 10 }, 2);
const g = boundGen(3);
__check(__line([...g].join(",")), "20,30");
