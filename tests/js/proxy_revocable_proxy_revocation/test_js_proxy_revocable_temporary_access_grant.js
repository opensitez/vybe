// vybe-test: js/proxy_revocable_proxy_revocation/test_js_proxy_revocable_temporary_access_grant
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

function withTemporaryAccess(target, fn) {
    const { proxy, revoke } = Proxy.revocable(target, {});
    try {
        return fn(proxy);
    } finally {
        revoke();
    }
}
const secretObj = { token: "XYZ-123" };
const result = withTemporaryAccess(secretObj, p => p.token);
__check(__line(result), "XYZ-123");
