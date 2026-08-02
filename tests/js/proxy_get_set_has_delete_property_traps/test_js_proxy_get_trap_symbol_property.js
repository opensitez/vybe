// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_get_trap_symbol_property
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

const sym = Symbol("test");
const target = { [sym]: "original" };
const proxy = new Proxy(target, {
    get(t, prop) {
        return typeof prop === "symbol" ? "intercepted_symbol" : t[prop];
    }
});
__check(__line(proxy[sym] + "|" + proxy.regular), "intercepted_symbol|undefined");
