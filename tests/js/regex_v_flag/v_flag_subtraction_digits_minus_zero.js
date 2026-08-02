// vybe-test: js/regex_v_flag/v_flag_subtraction_digits_minus_zero
// origin: languages/js/tests/js/test_regex_v_flag.rs

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

// Digits 1-9 (decimal digits minus zero)
const re = /^[1-9]+$/;
__check(__line(re.test("123456789")), "true");
__check(__line(re.test("1230")), "false");
