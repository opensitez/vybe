// vybe-test: js/regexp_sticky_y_and_global_g_flags/test_js_regexp_global_flag_advances_lastindex
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

const re = /\w+/g;
const str = "one two";
const m1 = re.exec(str);
const m2 = re.exec(str);
__check(__line(`${m1[0]}:${re.lastIndex}|${m2[0]}:${re.lastIndex}`), "one:3|two:7");
