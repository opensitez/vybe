// vybe-test: js/string_unicode_normalization/string_normalize_nfc_composes_characters
// origin: languages/js/tests/js/test_string_unicode_normalization.rs

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

const s="e\u0301"; __check(__line(s.normalize("NFC")=== "\u00E9"), "true");
