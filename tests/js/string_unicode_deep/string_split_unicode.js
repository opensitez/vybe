// vybe-test: js/string_unicode_deep/string_split_unicode
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

const s = "a,b,c";
const parts = s.split(",");
__check(__line(parts.length), "3");
__check(__line(parts.join("|")), "a|b|c");
