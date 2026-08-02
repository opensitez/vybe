// vybe-test: js/string_unicode/charcodeat_returns_code_unit_not_codepoint
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

const emoji = "😀";
const cu = emoji.charCodeAt(0);
const cp = emoji.codePointAt(0);
__check(__line(cu !== cp), "true");
__check(__line(cu === 0xD83D), "true");
