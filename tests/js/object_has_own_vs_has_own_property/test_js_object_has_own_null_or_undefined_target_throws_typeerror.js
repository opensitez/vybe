// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_null_or_undefined_target_throws_typeerror
// origin: languages/js/tests/js/test_js_object_has_own_vs_has_own_property.rs

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
    Object.hasOwn(null, "prop");
} catch (e) {
    __check(__line("hasOwn Null TypeError"), "hasOwn Null TypeError");
}
