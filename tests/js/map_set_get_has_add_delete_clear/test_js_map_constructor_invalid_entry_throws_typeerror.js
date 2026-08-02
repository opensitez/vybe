// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_constructor_invalid_entry_throws_typeerror
// origin: languages/js/tests/js/test_js_map_set_get_has_add_delete_clear.rs

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
    new Map(["not_a_tuple"]);
} catch (e) {
    __check(__line("Map Entry Non-Object TypeError"), "Map Entry Non-Object TypeError");
}
