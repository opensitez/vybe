// vybe-test: js/regex_string_methods/split_regex_with_capture_group
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

// Capture groups are included in result
const parts = "a1b2c".split(/(\d)/);
__check(__line(parts.join(",")), "a,1,b,2,c");
