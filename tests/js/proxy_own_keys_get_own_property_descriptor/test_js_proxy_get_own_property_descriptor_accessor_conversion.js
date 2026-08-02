// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_get_own_property_descriptor_accessor_conversion
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

const target = { count: 10 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return {
            get() { return 99; },
            enumerable: true,
            configurable: true
        };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "count");
__check(__line(typeof desc.get + "|" + proxy.count), "function|10");
