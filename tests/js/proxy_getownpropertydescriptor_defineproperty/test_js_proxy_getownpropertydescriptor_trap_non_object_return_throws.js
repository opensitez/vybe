// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_getownpropertydescriptor_trap_non_object_return_throws
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

const proxy = new Proxy({ a: 1 }, {
    getOwnPropertyDescriptor() { return "not_an_object"; }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "a");
} catch (e) {
    __check(__line("Descriptor Trap Non-Object TypeError"), "Descriptor Trap Non-Object TypeError");
}
