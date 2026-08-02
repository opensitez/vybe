// vybe-test: js/regex_patterns_deep/non_capturing_group
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

const re = /(?:foo)(bar)/;
const m = "foobar".match(re);
__check(__line(m[0]), "foobar"); // full match
__check(__line(m[1]), "bar"); // first capturing group (bar)
__check(__line(m[2]), "undefined"); // no second group
