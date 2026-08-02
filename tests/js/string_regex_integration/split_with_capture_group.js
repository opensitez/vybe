// vybe-test: js/string_regex_integration/split_with_capture_group
// origin: languages/js/tests/js/test_string_regex_integration.rs

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

// Capture groups appear in the split result
const parts = "one-two+three".split(/([-+])/);
__check(__line(parts.join(",")), "one,-,two,+,three");
