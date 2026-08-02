// vybe-test: js/string_legacy_edges/string_includes_empty_string_is_true_even_at_clamped_end
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

__check(__line("abc".includes("", 3)), "true");
__check(__line("abc".includes("", 99)), "true");
