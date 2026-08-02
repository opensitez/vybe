// vybe-test: js/regex_basics_matrix/regexp_whitespace_escape_matches_spaces
// origin: languages/js/tests/js/test_regex_basics_matrix.rs

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

__check(__line(/a\sb/.test("a b")), "true");
__check(__line(/a\sb/.test("ab")), "false");
