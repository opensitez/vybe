// vybe-test: js/regex_patterns_deep/positive_lookbehind
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

// Match digits preceded by "$"
const re = /(?<=\$)\d+/g;
const matches = [..."$100 €200 $300".matchAll(re)].map(m => m[0]);
__check(__line(matches.join(",")), "100,300");
