// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_getownpropertydescriptor_trap_basic
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

const target = { a: 10 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 99, writable: true, enumerable: true, configurable: true };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "a");
__check(__line(desc.value), "99");
