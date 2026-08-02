// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_returning_false_throws_typeerror_in_strict
// origin: languages/js/tests/js/test_js_proxy_getownpropertydescriptor_defineproperty.rs

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
    defineProperty() { return false; }
});
try {
    "use strict";
    Object.defineProperty(proxy, "a", { value: 1 });
} catch (e) {
    __check(__line("DefineProperty Trap Returned False TypeError"), "DefineProperty Trap Returned False TypeError");
}
