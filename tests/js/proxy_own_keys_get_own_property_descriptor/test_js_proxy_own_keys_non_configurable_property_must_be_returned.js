// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_non_configurable_property_must_be_returned
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
Object.defineProperty(target, "locked", { value: 10, configurable: false });
const proxy = new Proxy(target, {
    ownKeys(t) {
        return []; // Omitting non-configurable property violates invariant!
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    __check(__line("Non-Configurable OwnKeys Invariant Error"), "Non-Configurable OwnKeys Invariant Error");
}
