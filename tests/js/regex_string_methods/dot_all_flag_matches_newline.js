// vybe-test: js/regex_string_methods/dot_all_flag_matches_newline
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

const re = /a.b/s;
__check(__line(re.test("a\nb")), "true");
__check(__line(re.flags.includes("s")), "true");
