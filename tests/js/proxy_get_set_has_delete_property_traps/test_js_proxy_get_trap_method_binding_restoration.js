// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_get_trap_method_binding_restoration
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

const obj = {
    multiplier: 3,
    calc(x) { return x * this.multiplier; }
};
const proxy = new Proxy(obj, {
    get(t, prop, receiver) {
        const val = Reflect.get(t, prop, receiver);
        return typeof val === "function" ? val.bind(receiver) : val;
    }
});
const fn = proxy.calc;
__check(__line(fn(5)), "15");
