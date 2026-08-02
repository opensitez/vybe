// vybe-test: js/regex_modern_flags/regex_sticky_fails_if_match_does_not_start_at_lastindex
// origin: languages/js/tests/js/test_regex_modern_flags.rs

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

const re = /foo/y;
re.lastIndex = 0;
__check(__line(re.exec("bar foo") === null), "true");
__check(__line(re.lastIndex), "0");
