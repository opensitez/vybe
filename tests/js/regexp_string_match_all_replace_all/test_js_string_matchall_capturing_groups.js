// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_matchall_capturing_groups
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

const str = "a1 b2 c3";
const matches = [...str.matchAll(/([a-z])(\d)/g)];
__check(__line(matches.map(m => `${m[1]}:${m[2]}`).join(",")), "a:1,b:2,c:3");
