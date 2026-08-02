// vybe-test: js/regexp_unicode_sets_v_flag/test_js_regexp_unicode_sets_nested_character_classes
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

const re = /[[a-z]--[aeiou]]/v; // Consonants only
__check(__line(re.test("b") + "|" + re.test("e")), "true|false");
