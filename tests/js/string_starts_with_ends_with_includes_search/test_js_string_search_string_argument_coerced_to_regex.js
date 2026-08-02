// vybe-test: js/string_starts_with_ends_with_includes_search/test_js_string_search_string_argument_coerced_to_regex
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

const str = "hello 123 world";
__check(__line(str.search("123")), "6"); // String argument in search() is implicitly converted to new RegExp("123")!
