// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_trap_returns_non_array_throws
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

const proxy = new Proxy({}, {
    ownKeys() {
        return "not_an_array_or_object";
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    __check(__line("Non-List OwnKeys Error"), "Non-List OwnKeys Error");
}
