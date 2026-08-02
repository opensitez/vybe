// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_preventextensions_returns_false_throws_in_strict
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

const proxy = new Proxy({}, {
    preventExtensions() { return false; }
});
try {
    "use strict";
    Object.preventExtensions(proxy);
} catch (e) {
    __check(__line("preventExtensions Returned False TypeError"), "preventExtensions Returned False TypeError");
}
