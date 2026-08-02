// vybe-test: js/regex_flags_advanced/sticky_flag_anchors_at_last_index
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const re = /\d+/y;
re.lastIndex = 4;
const m = re.exec("abc 123 def");
__check(__line(m ? m[0] : null), "123");
__check(__line(re.lastIndex), "7");
