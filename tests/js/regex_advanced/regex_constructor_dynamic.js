// vybe-test: js/regex_advanced/regex_constructor_dynamic
// origin: languages/js/tests/js/test_regex_advanced.rs

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

function findWord(text, word) {
    let re = new RegExp("\\b" + word + "\\b", "gi");
    let matches = text.match(re);
    return matches ? matches.length : 0;
}
__check(__line(findWord("The the THE", "the")), "3");
