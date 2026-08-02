// vybe-test: js/functional_fp_patterns/partial_application
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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

function partial(fn, ...args) {
    return (...rest) => fn(...args, ...rest);
}
const add = (a, b, c) => a + b + c;
const add10 = partial(add, 10);
const add10and20 = partial(add, 10, 20);
__check(__line(add10(5, 3)), "18");
__check(__line(add10and20(7)), "37");
