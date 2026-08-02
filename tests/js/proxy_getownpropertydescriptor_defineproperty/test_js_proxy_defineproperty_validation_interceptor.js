// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_validation_interceptor
// origin: languages/js/tests/js/test_js_proxy_getownpropertydescriptor_defineproperty.rs

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
    defineProperty(t, prop, desc) {
        if (typeof desc.value !== "number") throw new TypeError("Must be number");
        return Reflect.defineProperty(t, prop, desc);
    }
});
proxy.age = 25;
__check(__line(proxy.age), "25");
try {
    proxy.age = "invalid";
} catch (e) {
    __check(__line(e.message), "Must be number");
}
