// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_basic_next_yield_values
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
    yield 1;
    yield 2;
    return 3;
}
const g = gen();
__check(__line(`${JSON.stringify(g.next())}:${JSON.stringify(g.next())}:${JSON.stringify(g.next())}`), "{\"value\":1,\"done\":false}:{\"value\":2,\"done\":false}:{\"value\":3,\"done\":true}");
