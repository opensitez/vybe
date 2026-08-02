// vybe-test: js/functional_patterns_deep/curry_creates_partial_applications
// origin: languages/js/tests/js/test_functional_patterns_deep.rs

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

const curry = fn => {
    const arity = fn.length;
    return function curried(...args) {
        if (args.length >= arity) return fn(...args);
        return (...more) => curried(...args, ...more);
    };
};
const add = curry((a, b, c) => a + b + c);
__check(__line(add(1)(2)(3)), "6");
__check(__line(add(1, 2)(3)), "6");
__check(__line(add(1)(2, 3)), "6");
