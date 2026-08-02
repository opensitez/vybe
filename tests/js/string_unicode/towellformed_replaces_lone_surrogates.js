// vybe-test: js/string_unicode/towellformed_replaces_lone_surrogates
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

// NFC normalization preserves emoji characters unchanged
const emoji = "😀";
const normalized = emoji.normalize("NFC");
__check(__line(normalized === emoji), "true");
__check(__line(normalized.length), "2");
