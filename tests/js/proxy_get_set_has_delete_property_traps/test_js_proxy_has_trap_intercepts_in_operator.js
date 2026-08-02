// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_has_trap_intercepts_in_operator
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

const target = { _secret: 42, public: 1 };
const proxy = new Proxy(target, {
    has(t, prop) {
        if (prop.startsWith("_")) return false;
        return prop in t;
    }
});
__check(__line(("_secret" in proxy) + "|" + ("public" in proxy)), "false|true");
