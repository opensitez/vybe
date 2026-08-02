// vybe-test: js/regexp_sticky_y_and_global_g_flags/test_js_regexp_sticky_zero_length_match_advances_lastindex
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

const re = /^|a/y;
__check(__line(re.test("a") + "|last=" + re.lastIndex), "true|last=0");
__check(__line(re.test("a") + "|last=" + re.lastIndex), "true|last=1");
