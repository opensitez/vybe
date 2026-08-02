// vybe-test: js/string_normalize_nfc_nfd_nfkc_nfkd/test_js_string_normalize_coerces_this_to_string
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

const res = String.prototype.normalize.call(12345, "NFC");
__check(__line(typeof res + "|" + res), "string|12345");
