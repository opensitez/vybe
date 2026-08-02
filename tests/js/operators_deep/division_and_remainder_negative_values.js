// vybe-test: js/operators_deep/division_and_remainder_negative_values
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

__check(__line(5 / -2), "-2.5");
__check(__line(-5 / 2), "-2.5");
__check(__line(-5 % 2), "-1");
__check(__line(5 % -2), "1");
__check(__line((-5 % -2)), "-1");
