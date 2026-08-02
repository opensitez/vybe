// vybe-test: js/regexp_lookbehind_assertions/test_js_regexp_lookbehind_unicode_flag
// origin: languages/js/tests/js/test_js_regexp_lookbehind_assertions.rs

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

const re = /(?<=\u{1F600})\w+/u;
const match = re.exec("😀happy");
__check(__line(match[0]), "happy");
