// vybe-test: js/string_unicode/normalize_nfkc_folds_compatibility
// origin: languages/js/tests/js/test_string_unicode.rs

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

const full = "\uFF41"; // fullwidth 'a'
const nfkc = full.normalize("NFKC");
__check(__line(nfkc === "a"), "true");
