// vybe-test: js/generator_protocol_advanced/generator_throw_causes_error_at_yield_point
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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
        yield 1;
    } catch (e) {
        yield "caught:" + e.message;
    }
}
const g = gen();
g.next(); // advance to yield 1
const result = g.throw(new Error("boom"));
__check(__line(result.value), "caught:boom");
__check(__line(result.done), "false");
