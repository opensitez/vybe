// vybe-test: js/ecma_functions/spread_in_call
// origin: languages/js/tests/js/test_ecma_functions.rs

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

function add(a, b, c) {
    return a + b + c;
}
const args = [1, 2, 3];
__check(__line(add(...args)), "6");
