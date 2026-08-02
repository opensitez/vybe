// vybe-test: js/string_unicode/from_codepoint_beyond_bmp
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

const s = String.fromCodePoint(0x10FFFF);
__check(__line(s.length), "2");
__check(__line(s.charCodeAt(0).toString(16)), "dbff"); // high surrogate of U+10FFFF
