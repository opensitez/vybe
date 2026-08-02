// vybe-test: js/math_extended/test_math_min_spread
// origin: languages/js/tests/js/test_math_extended.rs

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

let nums = [3, 1, 4, 1, 5, 9, 2, 6];
__check(__line(Math.min(...nums)), "1");
