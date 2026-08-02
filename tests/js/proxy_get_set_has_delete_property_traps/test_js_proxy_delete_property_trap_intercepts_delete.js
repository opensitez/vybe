// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_delete_property_trap_intercepts_delete
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

const target = { a: 1, protectedKey: 2 };
const proxy = new Proxy(target, {
    deleteProperty(t, prop) {
        if (prop === "protectedKey") return false; // Delete denied
        delete t[prop];
        return true;
    }
});
__check(__line(delete proxy.a), "true");
__check(__line(delete proxy.protectedKey), "false");
__check(__line(target.protectedKey), "2");
