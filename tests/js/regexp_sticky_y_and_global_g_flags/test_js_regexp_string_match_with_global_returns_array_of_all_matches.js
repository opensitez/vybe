// vybe-test: js/regexp_sticky_y_and_global_g_flags/test_js_regexp_string_match_with_global_returns_array_of_all_matches
// origin: languages/js/tests/js/test_js_regexp_sticky_y_and_global_g_flags.rs

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

const str = "test1 test2 test3";
const matches = str.match(/test\d/g);
__check(__line(matches.join(",")), "test1,test2,test3");
