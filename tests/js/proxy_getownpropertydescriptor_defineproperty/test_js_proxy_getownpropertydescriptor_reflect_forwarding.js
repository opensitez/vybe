// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_getownpropertydescriptor_reflect_forwarding
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

const target = { x: 50 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        const desc = Reflect.getOwnPropertyDescriptor(t, prop);
        desc.value *= 2;
        return desc;
    }
});
__check(__line(Object.getOwnPropertyDescriptor(proxy, "x").value), "100");
