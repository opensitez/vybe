// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_match_all_named_groups
// origin: languages/js/tests/js/test_js_regexp_named_capture_groups_indices.rs

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

const re = /(?<letter>[a-z])(?<digit>\d)/g;
const matches = [... "a1b2c3".matchAll(re)];
__check(__line(matches.map(m => `${m.groups.letter}:${m.groups.digit}`).join(",")), "a:1,b:2,c:3");
