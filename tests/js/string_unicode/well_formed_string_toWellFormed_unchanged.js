// vybe-test: js/string_unicode/well_formed_string_toWellFormed_unchanged
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

// normalize("NFC") on a plain ASCII string is a no-op
const s = "hello world";
__check(__line(s.normalize("NFC") === s), "true");
