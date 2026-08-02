// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_read_only_target_throws
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

const target = Object.freeze({ fixed: 1 });
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        return Reflect.defineProperty(t, prop, desc);
    }
});
try {
    "use strict";
    Object.defineProperty(proxy, "fixed", { value: 2 });
} catch (e) {
    __check(__line("DefineProperty Frozen Target TypeError"), "DefineProperty Frozen Target TypeError");
}
