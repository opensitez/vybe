// vybe-test: js/operators_deep/remainder_with_mixed_sign_and_float_values
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(10 % 4), "2");
__check(__line(10 % -4), "2");
__check(__line(-10 % 4), "-2");
__check(__line(10.5 % 2), "0.5");
