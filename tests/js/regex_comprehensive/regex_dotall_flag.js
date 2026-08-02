// vybe-test: js/regex_comprehensive/regex_dotall_flag
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

const text = "line1\nline2\nline3";
const withDot = text.match(/line1.line2/s);
const withoutDot = text.match(/line1.line2/);
__check(__line(withDot !== null), "true");
__check(__line(withoutDot), "null");
