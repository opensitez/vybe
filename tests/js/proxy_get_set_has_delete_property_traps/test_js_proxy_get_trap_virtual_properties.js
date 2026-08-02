// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_get_trap_virtual_properties
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

const proxy = new Proxy({}, {
    get(t, prop) {
        return `Virtual_${prop}`;
    }
});
__check(__line(proxy.foo + "|" + proxy.bar), "Virtual_foo|Virtual_bar");
