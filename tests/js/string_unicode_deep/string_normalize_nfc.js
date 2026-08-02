// vybe-test: js/string_unicode_deep/string_normalize_nfc
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

// é as combining: e + combining accent
const composed = "\u00E9"; // é precomposed
const decomposed = "e\u0301"; // e + combining accent
__check(__line(composed.length), "1");
__check(__line(decomposed.length), "2");
__check(__line(composed === decomposed), "false");
__check(__line(decomposed.normalize("NFC") === composed), "true");
