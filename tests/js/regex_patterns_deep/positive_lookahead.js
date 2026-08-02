// vybe-test: js/regex_patterns_deep/positive_lookahead
// origin: languages/js/tests/js/test_regex_patterns_deep.rs

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

// Match digits followed by "px"
const re = /\d+(?=px)/;
const m = "width: 100px, height: 200em".match(re);
__check(__line(m[0]), "100");
