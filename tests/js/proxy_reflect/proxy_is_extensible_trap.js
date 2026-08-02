// vybe-test: js/proxy_reflect/proxy_is_extensible_trap
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const obj = {};
const handler = {
    isExtensible(target) {
        // invariant: must match actual extensibility
        return Reflect.isExtensible(target);
    }
};
const p = new Proxy(obj, handler);
__check(__line(Object.isExtensible(p)), "true");
