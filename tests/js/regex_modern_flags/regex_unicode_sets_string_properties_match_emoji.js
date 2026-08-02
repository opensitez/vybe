// vybe-test: js/regex_modern_flags/regex_unicode_sets_string_properties_match_emoji
// origin: languages/js/tests/js/test_regex_modern_flags.rs

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

const re = /\p{RGI_Emoji}/v;
__check(__line(re.test("🙂")), "true");
__check(__line(re.test("A")), "false");
