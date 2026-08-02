// vybe-test: js/reflect_ownkeys_has_deleteproperty/test_js_reflect_has_non_object_target_throws_typeerror
// origin: languages/js/tests/js/test_js_reflect_ownkeys_has_deleteproperty.rs

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
    Reflect.has("not_an_object", "length");
} catch (e) {
    __check(__line("Reflect.has Non-Object TypeError"), "Reflect.has Non-Object TypeError");
}
