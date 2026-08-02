// vybe-test: js/string_starts_with_ends_with_includes_search/test_js_string_search_ignores_global_g_flag
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

const re = /a/g;
re.lastIndex = 5;
const pos = "cat bat".search(re);
__check(__line(pos + "|lastIndex=" + re.lastIndex), "1|lastIndex=5"); // search ignores g flag and does not mutate lastIndex!
