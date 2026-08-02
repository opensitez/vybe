// vybe-test: js/regex_patterns_deep/alternation_in_group
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

const re = /^(cat|dog|bird)$/;
__check(__line(re.test("cat")), "true");
__check(__line(re.test("dog")), "true");
__check(__line(re.test("fish")), "false");
