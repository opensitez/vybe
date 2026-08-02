// vybe-test: js/string_unicode_deep/string_normalize_nfd
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

const composed = "\u00E9";
const nfd = composed.normalize("NFD");
__check(__line(nfd.length), "2");
__check(__line(nfd.charCodeAt(0)), "101");
__check(__line(nfd.charCodeAt(1)), "769"); // combining acute accent
