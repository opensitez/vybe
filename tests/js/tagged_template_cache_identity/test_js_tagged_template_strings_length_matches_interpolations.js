// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_strings_length_matches_interpolations
// origin: languages/js/tests/js/test_js_tagged_template_cache_identity.rs

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

function tag(strings, ...values) {
    return strings.length === values.length + 1;
}
__check(__line(tag`A ${1} B ${2} C`), "true");
