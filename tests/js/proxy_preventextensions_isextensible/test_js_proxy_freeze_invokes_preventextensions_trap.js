// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_freeze_invokes_preventextensions_trap
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

let preventCalled = false;
const proxy = new Proxy({ a: 1 }, {
    preventExtensions(t) {
        preventCalled = true;
        return Reflect.preventExtensions(t);
    }
});
Object.freeze(proxy);
__check(__line(preventCalled + "|isFrozen=" + Object.isFrozen(proxy)), "true|isFrozen=true");
