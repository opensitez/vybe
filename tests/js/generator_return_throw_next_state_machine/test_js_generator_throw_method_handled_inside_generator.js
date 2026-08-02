// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_throw_method_handled_inside_generator
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
        yield 1;
    } catch (e) {
        yield "HandledInGen: " + e.message;
    }
}
const g = gen();
g.next();
__check(__line(g.throw(new Error("ExternalError")).value), "HandledInGen: ExternalError");
