// vybe-test: js/object_entries_keys_values_symbols/test_js_object_from_entries_invalid_pair_element_throws_typeerror
// origin: languages/js/tests/js/test_js_object_entries_keys_values_symbols.rs

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

try {
    Object.fromEntries([["valid", 1], "invalid_pair"]);
} catch (e) {
    __check(__line("fromEntries Invalid Pair TypeError"), "fromEntries Invalid Pair TypeError");
}
