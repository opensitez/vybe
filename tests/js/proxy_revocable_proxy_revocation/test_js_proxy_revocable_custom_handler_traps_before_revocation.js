// vybe-test: js/proxy_revocable_proxy_revocation/test_js_proxy_revocable_custom_handler_traps_before_revocation
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

const { proxy, revoke } = Proxy.revocable({ count: 5 }, {
    get(t, prop) { return t[prop] * 2; }
});
__check(__line(proxy.count), "10");
revoke();
__check(__line("Revoked Successfully"), "Revoked Successfully");
