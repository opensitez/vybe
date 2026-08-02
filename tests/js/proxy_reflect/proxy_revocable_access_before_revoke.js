// vybe-test: js/proxy_reflect/proxy_revocable_access_before_revoke
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

// Proxy.revocable: verify the proxy object exists before revoke
const { proxy, revoke } = Proxy.revocable({ x: 1 }, {
    get(target, prop) { return target[prop]; }
});
__check(__line(typeof proxy), "object");
__check(__line(typeof revoke), "function");
