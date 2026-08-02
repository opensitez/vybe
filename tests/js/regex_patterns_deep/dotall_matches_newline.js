// vybe-test: js/regex_patterns_deep/dotall_matches_newline
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

const re = /start.+end/s; // s flag = dotAll
const text = "start\nmiddle\nend";
__check(__line(re.test(text)), "true");
// Without s flag
const re2 = /start.+end/;
__check(__line(re2.test(text)), "false");
