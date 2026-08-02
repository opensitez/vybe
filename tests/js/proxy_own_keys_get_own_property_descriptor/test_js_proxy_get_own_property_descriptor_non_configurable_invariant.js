// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_get_own_property_descriptor_non_configurable_invariant
// origin: languages/js/tests/js/test_js_proxy_own_keys_get_own_property_descriptor.rs

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
Object.defineProperty(target, "locked", { value: 1, configurable: false });
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 1, configurable: true }; // Attempt to change configurable to true violates invariant!
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "locked");
} catch (e) {
    __check(__line("Descriptor Invariant Error"), "Descriptor Invariant Error");
}
