// vybe-test: js/regexp_sticky_y_and_global_g_flags/test_js_regexp_non_global_non_sticky_ignores_lastindex
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

const re = /a/;
re.lastIndex = 2;
const match = re.exec("cat");
__check(__line(match.index + "|lastIndex=" + re.lastIndex), "1|lastIndex=2");
