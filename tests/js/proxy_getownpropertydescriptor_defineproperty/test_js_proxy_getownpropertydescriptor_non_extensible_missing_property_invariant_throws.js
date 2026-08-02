// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_getownpropertydescriptor_non_extensible_missing_property_invariant_throws
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

const target = Object.preventExtensions({ a: 1 });
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 2, configurable: true, enumerable: true, writable: true }; // Invariant: Non-existent property on non-extensible target!
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "b");
} catch (e) {
    __check(__line("Non-Extensible Missing Property Descriptor TypeError"), "Non-Extensible Missing Property Descriptor TypeError");
}
