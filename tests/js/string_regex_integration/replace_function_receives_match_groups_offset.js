// vybe-test: js/string_regex_integration/replace_function_receives_match_groups_offset
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

const result = "2024-06-15".replace(
    /(\d{4})-(\d{2})-(\d{2})/,
    (full, year, month, day) => `${day}/${month}/${year}`
);
__check(__line(result), "15/06/2024");
