// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_coerces_boolean_return
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
        Reflect.defineProperty(t, prop, desc);
        return 1; // Truthy value 1 is coerced to boolean true
    }
});
Object.defineProperty(proxy, "val", { value: 10 });
__check(__line(proxy.val), "10");
