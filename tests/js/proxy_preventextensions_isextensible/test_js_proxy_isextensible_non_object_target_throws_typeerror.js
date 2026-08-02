// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_isextensible_non_object_target_throws_typeerror
// origin: languages/js/tests/js/test_js_proxy_preventextensions_isextensible.rs

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
    Reflect.isExtensible("not_an_object");
} catch (e) {
    __check(__line("Reflect.isExtensible Non-Object TypeError"), "Reflect.isExtensible Non-Object TypeError");
}
