// vybe-test: js/string_normalize_nfc_nfd_nfkc_nfkd/test_js_string_normalize_nfd_canonical_decomposition
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

const str = "\u00C5"; // Composed Å
const normalized = str.normalize("NFD");
__check(__line(normalized.length + "|c0=" + normalized.charCodeAt(0) + "|c1=" + normalized.charCodeAt(1)), "2|c0=65|c1=778");
