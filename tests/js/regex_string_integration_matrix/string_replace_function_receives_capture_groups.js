// vybe-test: js/regex_string_integration_matrix/string_replace_function_receives_capture_groups
// origin: languages/js/tests/js/test_regex_string_integration_matrix.rs

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

__check(__line("2024-07".replace(/(\d{4})-(\d{2})/, (_, y, m) => m + "/" + y)), "07/2024");
