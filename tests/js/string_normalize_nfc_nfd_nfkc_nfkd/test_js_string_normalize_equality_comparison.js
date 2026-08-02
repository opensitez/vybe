// vybe-test: js/string_normalize_nfc_nfd_nfkc_nfkd/test_js_string_normalize_equality_comparison
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

const s1 = "\u00C9"; // É composed
const s2 = "\u0045\u0301"; // E + combining acute
__check(__line((s1 === s2) + "|" + (s1.normalize() === s2.normalize())), "false|true");
