// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_non_configurable_non_existent_invariant_throws
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
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        return true; // Trap reports success without actually defining non-configurable property on target!
    }
});
try {
    Object.defineProperty(proxy, "a", { value: 1, configurable: false });
} catch (e) {
    __check(__line("DefineProperty Non-Configurable Invariant TypeError"), "DefineProperty Non-Configurable Invariant TypeError");
}
