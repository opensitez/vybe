// vybe-test: js/string_methods_deep/replace_replaces_first_match_only
// origin: languages/js/tests/js/test_string_methods_deep.rs

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

__check(__line("a-b-b".replace("-", "_")), "a_b-b");
__check(__line("a-b-b".replace(/-/g, "_")), "a_b_b");
