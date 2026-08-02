// vybe-test: js/regex_comprehensive/regex_sticky_flag
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

const re = /\d+/y;
re.lastIndex = 3;
const str = "abc123def456";
const m1 = re.exec(str);
__check(__line(m1[0]), "123");
__check(__line(re.lastIndex), "6");
const m2 = re.exec(str);
__check(__line(m2), "null");
