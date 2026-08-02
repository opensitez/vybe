// vybe-test: js/ecma_operators/arithmetic_nan_and_infinity
// origin: languages/js/tests/js/test_ecma_operators.rs

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

__check(__line(1 / 0), "Infinity");
__check(__line(-1 / 0), "-Infinity");
__check(__line(0 / 0), "NaN");
__check(__line(10 + NaN), "NaN");
__check(__line(Infinity - Infinity), "NaN");
