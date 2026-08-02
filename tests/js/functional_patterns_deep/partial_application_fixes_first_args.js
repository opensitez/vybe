// vybe-test: js/functional_patterns_deep/partial_application_fixes_first_args
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

function partial(fn, ...preset) {
    return (...rest) => fn(...preset, ...rest);
}
function multiply(a, b, c) { return a * b * c; }
const double = partial(multiply, 2);
const triple = partial(multiply, 3);
__check(__line(double(3, 4)), "24");
__check(__line(triple(2, 5)), "30");
