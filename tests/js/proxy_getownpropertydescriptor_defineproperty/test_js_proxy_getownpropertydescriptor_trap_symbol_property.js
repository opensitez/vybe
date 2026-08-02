// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_getownpropertydescriptor_trap_symbol_property
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

const sym = Symbol("id");
const proxy = new Proxy({}, {
    getOwnPropertyDescriptor(t, prop) {
        if (prop === sym) return { value: "SymValue", configurable: true, enumerable: true };
    }
});
__check(__line(Object.getOwnPropertyDescriptor(proxy, sym).value), "SymValue");
