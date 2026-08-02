// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_getownpropertydescriptor_non_configurable_invariant_violation_throws
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

const target = {};
Object.defineProperty(target, "fixed", { value: 1, configurable: false });
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 1, configurable: true }; // Invariant: Cannot report non-configurable property as configurable!
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "fixed");
} catch (e) {
    __check(__line("Descriptor Invariant TypeError"), "Descriptor Invariant TypeError");
}
