// vybe-test: js/generators/generator_throw_before_first_yield_caught
// origin: languages/js/tests/js/test_generators.rs

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

function* guarded() {
    try {
        yield "ready";
    } catch (err) {
        __check(__line("caught: " + err.message), "caught: stop");
        yield "handled";
    }
}
let g = guarded();
let result = g.throw(new Error("stop"));
__check(__line(result.value), "handled");
__check(__line(result.done), "false");
