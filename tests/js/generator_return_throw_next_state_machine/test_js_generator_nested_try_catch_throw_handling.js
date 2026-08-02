// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_nested_try_catch_throw_handling
// origin: languages/js/tests/js/test_js_generator_return_throw_next_state_machine.rs

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
    try {
        try {
            yield "inner";
        } catch (e) {
            yield "caughtInner: " + e.message;
            throw new Error("outerErr");
        }
    } catch (e) {
        yield "caughtOuter: " + e.message;
    }
}
const g = gen();
g.next();
g.throw(new Error("initial"));
__check(__line(g.next().value), "caughtOuter: outerErr");
