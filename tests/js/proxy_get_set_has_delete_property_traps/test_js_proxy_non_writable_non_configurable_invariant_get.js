// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_non_writable_non_configurable_invariant_get
// origin: languages/js/tests/js/test_js_proxy_get_set_has_delete_property_traps.rs

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
Object.defineProperty(target, "fixed", {
    value: 100,
    writable: false,
    configurable: false
});
const proxy = new Proxy(target, {
    get(t, prop) {
        return 999; // Attempt to violate invariant!
    }
});
try {
    proxy.fixed;
} catch (e) {
    __check(__line("Proxy Invariant Get Violation"), "Proxy Invariant Get Violation");
}
