// vybe-test: js/regex_patterns_deep/multiline_mode_anchors
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

const text = "first\nsecond\nthird";
const matches = text.match(/^\w+/mg);
__check(__line(matches.join(",")), "first,second,third");
