// vybe-test: js/regex_basics_matrix/regexp_global_exec_no_more_matches_resets_lastindex
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

const re = /a/g;
re.exec("a");
__check(__line(re.exec("a") === null), "true");
__check(__line(re.lastIndex), "0");
