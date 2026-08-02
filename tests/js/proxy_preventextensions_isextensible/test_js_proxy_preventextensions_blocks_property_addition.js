// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_preventextensions_blocks_property_addition
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
    preventExtensions(t) {
        return Reflect.preventExtensions(t);
    }
});
Object.preventExtensions(proxy);
try {
    "use strict";
    proxy.newProp = 100;
} catch (e) {
    __check(__line("Add Property To Non-Extensible Proxy TypeError"), "Add Property To Non-Extensible Proxy TypeError");
}
