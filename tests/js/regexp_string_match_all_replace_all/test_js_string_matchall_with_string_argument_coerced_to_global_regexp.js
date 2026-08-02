// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_matchall_with_string_argument_coerced_to_global_regexp
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

const matches = [..."hello world".matchAll("l")];
__check(__line(matches.length + "|" + matches.map(m => m.index).join(",")), "3|2,3,9");
