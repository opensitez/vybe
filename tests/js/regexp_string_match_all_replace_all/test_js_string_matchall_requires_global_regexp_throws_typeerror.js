// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_matchall_requires_global_regexp_throws_typeerror
// origin: languages/js/tests/js/test_js_regexp_string_match_all_replace_all.rs

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
    "hello".matchAll(/l/); // Non-global RegExp throws TypeError!
} catch (e) {
    __check(__line("matchAll Non-Global TypeError"), "matchAll Non-Global TypeError");
}
