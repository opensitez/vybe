// vybe-test: js/string_unicode_deep/code_point_at_surrogate_pair
// origin: languages/js/tests/js/test_string_unicode_deep.rs

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

// 𝌆 is U+1D306, a supplemental character stored as a surrogate pair
const s = "𝌆";
__check(__line(s.length), "2"); // 2 UTF-16 code units
__check(__line(s.charCodeAt(0).toString(16)), "d834"); // "d834" — high surrogate
