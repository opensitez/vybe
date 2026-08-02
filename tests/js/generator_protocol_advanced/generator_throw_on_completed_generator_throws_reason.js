// vybe-test: js/generator_protocol_advanced/generator_throw_on_completed_generator_throws_reason
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

function* gen() { yield 1; }
const g = gen();
g.next();
g.next();
try {
    g.throw("custom_err");
} catch (e) {
    __check(__line(e), "custom_err");
}
