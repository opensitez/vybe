// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_nested_proxy_traps
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

const target = { value: 1 };
const proxy1 = new Proxy(target, {
    get(t, prop) { return t[prop] + 10; }
});
const proxy2 = new Proxy(proxy1, {
    get(t, prop) { return t[prop] * 2; }
});
__check(__line(proxy2.value), "22");
