// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_isextensible_trap_this_binding
// origin: languages/js/tests/js/test_js_proxy_preventextensions_isextensible.rs

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

let trapThis;
const handler = {
    isExtensible(t) {
        trapThis = this;
        return Reflect.isExtensible(t);
    }
};
const proxy = new Proxy({}, handler);
Object.isExtensible(proxy);
__check(__line(trapThis === handler), "true");
