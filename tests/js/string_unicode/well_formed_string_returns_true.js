// vybe-test: js/string_unicode/well_formed_string_returns_true
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

// BMP strings: spread length equals code-unit length
// Supplemental strings: spread length < code-unit length
__check(__line([..."hello"].length === "hello".length), "true");
__check(__line([..."😀"].length < "😀".length), "true");
