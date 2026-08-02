// vybe-test: js/string_normalize_nfc_nfd_nfkc_nfkd/test_js_string_normalize_null_or_undefined_this_throws_typeerror
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

try {
    String.prototype.normalize.call(null);
} catch (e) {
    __check(__line("normalize Null This TypeError"), "normalize Null This TypeError");
}
