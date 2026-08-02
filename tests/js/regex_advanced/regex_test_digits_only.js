// vybe-test: js/regex_advanced/regex_test_digits_only
// origin: languages/js/tests/js/test_regex_advanced.rs

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

let re = /^\d+$/;
__check(__line(re.test("12345")), "true");
__check(__line(re.test("123a5")), "false");
__check(__line(re.test("")), "false");
