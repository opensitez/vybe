// vybe-test: js/generator_protocol_advanced/generator_first_next_arg_is_always_ignored
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
    const x = yield 1;
    yield x;
}
const g = gen();
g.next("ignored"); // first next arg always ignored
const r = g.next(42);
__check(__line(r.value), "42"); // 42 — second next value received
