// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_set_trap_receiver_prototype_chain
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

const proto = new Proxy({}, {
    set(t, prop, val, receiver) {
        receiver["store_" + prop] = val;
        return true;
    }
});
const child = Object.create(proto);
child.field = 50;
__check(__line(child.store_field), "50");
