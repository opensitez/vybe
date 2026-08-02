// vybe-test: js/regex_string_integration_matrix/string_matchall_exposes_capture_groups
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

const values = [..."2024-07 2025-08".matchAll(/(\d{4})-(\d{2})/g)].map(m => m[1] + "/" + m[2]);
__check(__line(values.join(",")), "2024/07,2025/08");
