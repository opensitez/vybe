// vybe-test: js/string_starts_with_ends_with_includes_search/test_js_string_inspection_null_or_undefined_this_throws_typeerror
// origin: languages/js/tests/js/test_js_string_starts_with_ends_with_includes_search.rs

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
    String.prototype.includes.call(null, "test");
} catch (e) {
    __check(__line("includes Null This TypeError"), "includes Null This TypeError");
}
