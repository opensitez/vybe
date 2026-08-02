// vybe-test: js/closures_functional/curry_basic
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

function curry(fn) {
    return function(a) {
        return function(b) {
            return fn(a, b);
        };
    };
}
let add = curry((a, b) => a + b);
__check(__line(add(3)(4)), "7");
