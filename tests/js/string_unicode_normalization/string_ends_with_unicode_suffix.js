// vybe-test: js/string_unicode_normalization/string_ends_with_unicode_suffix
// origin: languages/js/tests/js/test_string_unicode_normalization.rs

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

__check(__line("naïve".endsWith("ïve")), "true");
