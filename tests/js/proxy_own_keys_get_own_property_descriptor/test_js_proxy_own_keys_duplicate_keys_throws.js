// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_duplicate_keys_throws
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
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["a", "a"]; // Duplicate keys not allowed in ownKeys result!
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    __check(__line("Duplicate OwnKeys Error"), "Duplicate OwnKeys Error");
}
