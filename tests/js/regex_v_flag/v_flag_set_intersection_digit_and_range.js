// vybe-test: js/regex_v_flag/v_flag_set_intersection_digit_and_range
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

// Digits 5-9 only
const re = /^[5-9]+$/;
__check(__line(re.test("579")), "true");
__check(__line(re.test("1234")), "false");
