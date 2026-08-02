// vybe-test: js/symbol_is_concat_spreadable_to_string_tag/test_js_symbol_tostringtag_non_string_ignored
// origin: languages/js/tests/js/test_js_symbol_is_concat_spreadable_to_string_tag.rs

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

const obj = {
    [Symbol.toStringTag]: 12345 // Non-string is ignored, falls back to default Object tag!
};
__check(__line(Object.prototype.toString.call(obj)), "[object Object]");
