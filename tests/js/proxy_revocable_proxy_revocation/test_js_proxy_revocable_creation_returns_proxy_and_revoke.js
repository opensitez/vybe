// vybe-test: js/proxy_revocable_proxy_revocation/test_js_proxy_revocable_creation_returns_proxy_and_revoke
// origin: languages/js/tests/js/test_js_proxy_revocable_proxy_revocation.rs

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

const { proxy, revoke } = Proxy.revocable({ a: 1 }, {});
__check(__line(proxy.a), "1");
__check(__line(typeof revoke), "function");
