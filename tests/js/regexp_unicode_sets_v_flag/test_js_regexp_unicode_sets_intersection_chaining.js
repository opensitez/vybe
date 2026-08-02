// vybe-test: js/regexp_unicode_sets_v_flag/test_js_regexp_unicode_sets_intersection_chaining
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

const re = /[a-z]&&[c-z]&&[a-f]/v; // Elements in a-z AND c-z AND a-f -> c,d,e,f
__check(__line(re.test("d") + "|" + re.test("a") + "|" + re.test("z")), "true|false|false");
