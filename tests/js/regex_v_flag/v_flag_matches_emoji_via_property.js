// vybe-test: js/regex_v_flag/v_flag_matches_emoji_via_property
// origin: languages/js/tests/js/test_regex_v_flag.rs

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

// Test emoji detection via code point range (emoji region starts at U+1F600)
const re = /[\u{1F600}-\u{1F64F}]/u;
__check(__line(re.test("😀")), "true");
__check(__line(re.test("abc")), "false");
