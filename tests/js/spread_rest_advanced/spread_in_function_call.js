// vybe-test: js/spread_rest_advanced/spread_in_function_call
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

function sum(a, b, c) { return a + b + c; }
const args = [1, 2, 3];
__check(__line(sum(...args)), "6");
