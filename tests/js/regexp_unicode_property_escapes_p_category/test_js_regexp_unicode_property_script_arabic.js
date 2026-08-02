// vybe-test: js/regexp_unicode_property_escapes_p_category/test_js_regexp_unicode_property_script_arabic
// origin: languages/js/tests/js/test_js_regexp_unicode_property_escapes_p_category.rs

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

const re = /\p{Script=Arabic}/gu;
__check(__line("مرحبا World".match(re).join("")), "مرحبا");
