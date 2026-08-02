// vybe-test: js/regex_comprehensive/regex_split_with_capture
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

const parts = "one1two22three333four".split(/(\d+)/);
__check(__line(parts.join("|")), "one|1|two|22|three|333|four");
