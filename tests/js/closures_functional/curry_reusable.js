// vybe-test: js/closures_functional/curry_reusable
// origin: languages/js/tests/js/test_closures_functional.rs

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

function multiply(a) {
    return function(b) {
        return a * b;
    };
}
let double = multiply(2);
let triple = multiply(3);
__check(__line(double(5)), "10");
__check(__line(triple(5)), "15");
