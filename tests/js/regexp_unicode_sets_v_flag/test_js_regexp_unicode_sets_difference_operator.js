// vybe-test: js/regexp_unicode_sets_v_flag/test_js_regexp_unicode_sets_difference_operator
// origin: languages/js/tests/js/test_js_regexp_unicode_sets_v_flag.rs

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

const re = /[\p{Decimal_Number}--[0-4]]/v; // Digits except 0..4
__check(__line(re.test("5") + "|" + re.test("3")), "true|false");
