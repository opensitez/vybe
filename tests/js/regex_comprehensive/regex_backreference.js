// vybe-test: js/regex_comprehensive/regex_backreference
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

// Match doubled words
const doubled = /\b(\w+) \1\b/g;
const text = "the the quick brown fox fox";
const matches = text.match(doubled);
__check(__line(matches.join(",")), "the the,fox fox");
