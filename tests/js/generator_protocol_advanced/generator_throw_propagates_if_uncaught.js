// vybe-test: js/generator_protocol_advanced/generator_throw_propagates_if_uncaught
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
    yield 1;
    yield 2;
}
const g = gen();
g.next();
let threw = false;
try {
    g.throw(new Error("err"));
} catch (e) {
    threw = true;
}
__check(__line(threw), "true");
