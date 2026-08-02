// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_passing_arguments_into_next
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
    const a = yield "first";
    const b = yield a * 2;
    return b + 10;
}
const g = gen();
__check(__line(g.next().value), "first"); // "first"
__check(__line(g.next(5).value), "10"); // 5 * 2 = 10
__check(__line(g.next(20).value), "30"); // 20 + 10 = 30
