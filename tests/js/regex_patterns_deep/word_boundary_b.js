// vybe-test: js/regex_patterns_deep/word_boundary_b
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

const re = /\bcat\b/;
__check(__line(re.test("cat")), "true");
__check(__line(re.test("cats")), "false");
__check(__line(re.test("the cat sat")), "true");
__check(__line(re.test("concatenate")), "false");
