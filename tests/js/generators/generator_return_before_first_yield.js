// vybe-test: js/generators/generator_return_before_first_yield
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

function* gen() {
    yield 1;
    return 99;
}
let g = gen();
let result = g.return("stopped");
__check(__line(result.value), "stopped");
__check(__line(result.done), "true");
__check(__line(g.next().done), "true");
