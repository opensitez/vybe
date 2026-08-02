// vybe-test: js/generators/spread_on_custom_iterable
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
    yield 10;
    yield 20;
    yield 30;
}
let arr = [...gen()];
__check(__line(arr.join(",")), "10,20,30");
