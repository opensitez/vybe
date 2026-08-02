// vybe-test: js/regex_string_integration_matrix/string_matchall_exposes_match_indexes
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

const values = [..."a1b22c333".matchAll(/\d+/g)].map(m => m.index);
__check(__line(values.join(",")), "1,3,6");
