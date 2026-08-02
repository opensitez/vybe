// vybe-test: js/regex_patterns_deep/negative_lookahead
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

// Match digits NOT followed by "px"
const re = /\d+(?!px)/g;
const matches = [..."200px 300em 400px".matchAll(re)].map(m => m[0]);
__check(__line(matches.join(",")), "20,300,40");
