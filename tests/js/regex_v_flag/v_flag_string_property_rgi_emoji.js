// vybe-test: js/regex_v_flag/v_flag_string_property_rgi_emoji
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

// Match common emoji range U+1F600-U+1F64F
const re = /[\u{1F600}-\u{1F64F}]/u;
__check(__line(re.test("😀")), "true");
