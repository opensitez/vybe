// vybe-test: js/regex_string_integration_matrix/string_replace_with_dollar_ampersand_uses_whole_match
// origin: languages/js/tests/js/test_regex_string_integration_matrix.rs

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

__check(__line("abc123def".replace(/\d+/, "[$&]")), "abc[123]def");
