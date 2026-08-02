// vybe-test: js/regex_flags_advanced/regex_lastindex_manual_control
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const re = /a/g;
const str = "ababa";
re.lastIndex = 2;
const m = re.exec(str);
__check(__line(m.index), "2"); // finds 'a' at index 2
re.lastIndex = 0;
const m2 = re.exec(str);
__check(__line(m2.index), "0"); // starts from beginning
