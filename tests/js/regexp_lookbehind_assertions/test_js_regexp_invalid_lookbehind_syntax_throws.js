// vybe-test: js/regexp_lookbehind_assertions/test_js_regexp_invalid_lookbehind_syntax_throws
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

try {
    eval("const re = /(?<=);/");
} catch (e) {
    __check(__line("Empty Lookbehind SyntaxError"), "Empty Lookbehind SyntaxError");
}
