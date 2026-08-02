// vybe-test: js/string_methods_more/string_normalize_nfc_canonicalizes_accents
// origin: languages/js/tests/js/test_string_methods_more.rs

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

const decomposed = "e\u0301";
const normalized = decomposed.normalize("NFC");
__check(__line(decomposed.length), "2");
__check(__line(normalized.length), "1");
__check(__line(normalized === "\u00e9"), "true");
