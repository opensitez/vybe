// vybe-test: js/regex_string_methods/dot_does_not_match_newline_by_default
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const re = /a.b/;
__check(__line(re.test("a\nb")), "false");
__check(__line(re.test("acb")), "true");
