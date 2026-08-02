// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_get_own_property_descriptor_trap
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

const target = { val: 42 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return {
            value: t[prop] * 2,
            writable: true,
            enumerable: true,
            configurable: true
        };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "val");
__check(__line(desc.value), "84");
