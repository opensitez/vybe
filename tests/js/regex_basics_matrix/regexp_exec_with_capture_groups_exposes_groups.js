// vybe-test: js/regex_basics_matrix/regexp_exec_with_capture_groups_exposes_groups
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

const m = /(ab)(cd)/.exec("xabcdz");
__check(__line(m[1]), "ab");
__check(__line(m[2]), "cd");
