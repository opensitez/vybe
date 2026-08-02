// vybe-test: js/string_regex_integration/match_without_global_returns_first_with_groups
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

const m = "2024-06-15".match(/(\d{4})-(\d{2})-(\d{2})/);
__check(__line(m[0]), "2024-06-15");
__check(__line(m[1]), "2024");
__check(__line(m[2]), "06");
