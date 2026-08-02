// vybe-test: js/string_unicode/spread_counts_codepoints_not_code_units
// origin: languages/js/tests/js/test_string_unicode.rs

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

const s = "a😀b";
const chars = [...s];
__check(__line(chars.length), "3");
__check(__line(s.length), "4");
