// vybe-test: js/regexp_sticky_y_and_global_g_flags/test_js_regexp_lastindex_read_only_when_frozen_throws
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

const re = /a/g;
Object.freeze(re);
try {
    "use strict";
    re.exec("a");
} catch (e) {
    __check(__line("Frozen RegExp Exec TypeError"), "Frozen RegExp Exec TypeError");
}
