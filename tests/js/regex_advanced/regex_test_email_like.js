// vybe-test: js/regex_advanced/regex_test_email_like
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

let re = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
__check(__line(re.test("user@example.com")), "true");
__check(__line(re.test("bad@")), "false");
__check(__line(re.test("test.user@domain.co.uk")), "true");
