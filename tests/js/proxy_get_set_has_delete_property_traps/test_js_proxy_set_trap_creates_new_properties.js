// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_set_trap_creates_new_properties
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
const proxy = new Proxy(target, {
    set(t, prop, val) {
        t["prefix_" + prop] = val;
        return true;
    }
});
proxy.data = 100;
__check(__line(target.prefix_data), "100");
