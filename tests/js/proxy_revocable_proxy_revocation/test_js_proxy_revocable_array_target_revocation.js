// vybe-test: js/proxy_revocable_proxy_revocation/test_js_proxy_revocable_array_target_revocation
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

const { proxy, revoke } = Proxy.revocable([10, 20, 30], {});
__check(__line(proxy[0]), "10");
revoke();
try {
    proxy.push(40);
} catch (e) {
    __check(__line("Array Push Revoked Error"), "Array Push Revoked Error");
}
