// vybe-test: js/generators_lazy/generator_body_does_not_eagerly_execute
// origin: languages/js/tests/js/test_generators_lazy.rs

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

function* loud() {
    console.log("bad: body ran before resume");
    yield 1;
}
let g = loud();
console.log("ok");
