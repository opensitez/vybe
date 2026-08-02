// vybe-test: js/regex_basics_matrix/string_match_without_global_returns_match_object
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

const m = "abc123".match(/\d+/);
__check(__line(m[0]), "123");
__check(__line(m.index), "3");
