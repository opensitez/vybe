// vybe-test: js/regex_string_methods/match_with_global_returns_all_matches
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

const result = "cat bat sat".match(/[a-z]at/g);
__check(__line(result.join(",")), "cat,bat,sat");
