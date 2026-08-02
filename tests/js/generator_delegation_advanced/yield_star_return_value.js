// vybe-test: js/generator_delegation_advanced/yield_star_return_value
// origin: languages/js/tests/js/test_generator_delegation_advanced.rs

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

function* sub() {
    yield 1;
    yield 2;
    return "done";
}
function* main() {
    const result = yield* sub();
    yield result;
}
__check(__line([...main()].join(",")), "1,2,done");
