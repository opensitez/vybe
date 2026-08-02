// vybe-test: js/regex_string_integration_matrix/regexp_source_escapes_forward_slash
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

__check(__line(new RegExp("a/b").source), "a\\/b");
