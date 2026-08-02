// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_return_method_forces_completion
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
    yield 10;
    yield 20;
}
const g = gen();
g.next();
const ret = g.return("ForcedReturn");
__check(__line(`${ret.value}:${ret.done}:${g.next().done}`), "ForcedReturn:true:true");
