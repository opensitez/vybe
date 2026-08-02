// vybe-test: js/string_normalize_nfc_nfd_nfkc_nfkd/test_js_string_normalize_nfd_combining_character_ordering
// origin: languages/js/tests/js/test_js_string_normalize_nfc_nfd_nfkc_nfkd.rs

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

const str = "e\u0301\u0327"; // e + acute + cedilla
const normalized = str.normalize("NFD");
__check(__line(normalized.length), "3");
