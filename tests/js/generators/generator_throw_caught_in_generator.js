// vybe-test: js/generators/generator_throw_caught_in_generator
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
        yield "after";
    } catch (err) {
        __check(__line("caught: " + err.message), "ready");
    }
}
let g = guarded();
__check(__line(g.next().value), "caught: stop");
let result = g.throw(new Error("stop"));
__check(__line(result.done), "true");
