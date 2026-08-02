// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_non_extensible_target_must_include_all_keys
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

const target = { x: 1, y: 2 };
Object.preventExtensions(target);
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["x"]; // Missing 'y' violates invariant for non-extensible target!
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    console.log("Non-Extensible OwnKeys Invariant Error");
}
