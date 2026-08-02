// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_throw_on_unstarted_generator
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

function* gen() { yield 1; }
const g = gen();
try {
    g.throw(new Error("UnstartedThrow"));
} catch (e) {
    __check(__line(e.message + "|done=" + g.next().done), "UnstartedThrow|done=true");
}
