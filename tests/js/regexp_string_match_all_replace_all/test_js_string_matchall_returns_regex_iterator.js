// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_matchall_returns_regex_iterator
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

const re = /t(e)(st(\d?))/g;
const str = "test1test2";
const matches = [...str.matchAll(re)];
__check(__line(matches.length + "|" + matches[0][0] + "|" + matches[1][0]), "2|test1|test2");
