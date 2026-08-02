// vybe-test: js/string_legacy_edges/string_repeat_with_fractional_count_uses_integer_part
// origin: languages/js/tests/js/test_string_legacy_edges.rs

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

__check(__line("ab".repeat(2.7)), "abab");
