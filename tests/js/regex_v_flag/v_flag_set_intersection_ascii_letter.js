// vybe-test: js/regex_v_flag/v_flag_set_intersection_ascii_letter
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

// ASCII letters only (intersection of Letter and ASCII)
const re = /^[a-zA-Z]+$/;
__check(__line(re.test("hello")), "true");
__check(__line(re.test("café")), "false");
