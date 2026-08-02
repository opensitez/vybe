// vybe-test: js/destructuring_patterns/function_param_destructure_array
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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

function sum([a, b, c = 0]) {
    return a + b + c;
}
__check(__line(sum([1, 2, 3])), "6");
__check(__line(sum([4, 5])), "9");
