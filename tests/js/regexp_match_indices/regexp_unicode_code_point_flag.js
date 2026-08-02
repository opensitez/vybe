// vybe-test: js/regexp_match_indices/regexp_unicode_code_point_flag
// origin: languages/js/tests/js/test_regexp_match_indices.rs

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

__check(__line(/👍/u.test("👍")), "true");
