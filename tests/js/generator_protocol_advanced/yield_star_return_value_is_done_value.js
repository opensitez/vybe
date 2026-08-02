// vybe-test: js/generator_protocol_advanced/yield_star_return_value_is_done_value
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
    const result = yield* (function*() {
        yield 1; yield 2;
        return "final";
    })();
    yield result; // "final" from delegated generator's done value
}
console.log([...gen()].join(","));
