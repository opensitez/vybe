// vybe-test: js/regex_basics_matrix/regexp_noncapturing_group_does_not_create_capture_slot
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

const m = /(?:ab)(cd)/.exec("abcd");
__check(__line(m.length), "2");
__check(__line(m[1]), "cd");
