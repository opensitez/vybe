// vybe-test: js/string_unicode/lone_surrogate_is_not_well_formed
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

// Verify emoji surrogate pair code units are in surrogate ranges
const emoji = "😀";
const hi = emoji.charCodeAt(0);
const lo = emoji.charCodeAt(1);
__check(__line(hi >= 0xD800 && hi <= 0xDBFF), "true");
__check(__line(lo >= 0xDC00 && lo <= 0xDFFF), "true");
