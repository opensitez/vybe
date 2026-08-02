// vybe-test: js/generators/destructure_generator
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
    yield "a";
    yield "b";
    yield "c";
}
let [x, y, z] = gen();
__check(__line(x), "a");
__check(__line(y), "b");
__check(__line(z), "c");
