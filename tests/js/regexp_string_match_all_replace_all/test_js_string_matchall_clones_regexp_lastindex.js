// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_matchall_clones_regexp_lastindex
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

const re = /a/g;
re.lastIndex = 2; // Original regex lastIndex should NOT affect matchAll or be modified by it!
const matches = [..."aaa".matchAll(re)];
__check(__line(matches.length + "|originalLastIndex=" + re.lastIndex), "3|originalLastIndex=2");
