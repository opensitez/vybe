// vybe-test: js/regex_string_methods/replace_with_named_capture_in_template
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

const result = "2024-01-15".replace(
    /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/,
    "$<month>/$<day>/$<year>"
);
__check(__line(result), "01/15/2024");
