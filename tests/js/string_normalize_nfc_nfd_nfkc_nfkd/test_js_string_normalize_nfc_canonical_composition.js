// vybe-test: js/string_normalize_nfc_nfd_nfkc_nfkd/test_js_string_normalize_nfc_canonical_composition
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

const str = "\u0041\u030A"; // Decomposed Å
const normalized = str.normalize("NFC");
__check(__line(normalized + "|len=" + normalized.length + "|code=" + normalized.charCodeAt(0)), "Å|len=1|code=197");
