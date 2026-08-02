// vybe-test: js/generators/generator_next_on_exhausted_generator_returns_done_true
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

function* g() { yield 1; }
const gen = g();
gen.next();
gen.next();
const r = gen.next();
__check(__line(r.value === undefined && r.done === true), "true");
