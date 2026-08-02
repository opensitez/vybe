// vybe-test: js/symbol_is_concat_spreadable_to_string_tag/test_js_symbol_is_concat_spreadable_defaults
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

const array = [1];
const plainObj = { 0: "a", length: 1 };
const res = [0].concat(array, plainObj);
__check(__line(res.length + "|" + Array.isArray(res[1]) + "|" + (typeof res[2])), "3|false|object");
