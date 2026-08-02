// vybe-test: js/string_unicode_deep/code_point_at_bmp
// origin: languages/js/tests/js/test_string_unicode_deep.rs

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

const s = "ABC";
__check(__line(s.codePointAt(0)), "65");
__check(__line(s.codePointAt(1)), "66");
