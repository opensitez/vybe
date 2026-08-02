// vybe-test: js/proxy_revocable_proxy_revocation/test_js_proxy_revocable_apply_trap_after_revoke_throws_typeerror
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

function fn() { return 100; }
const { proxy, revoke } = Proxy.revocable(fn, {});
revoke();
try {
    proxy();
} catch (e) {
    __check(__line("Revoked Proxy Apply TypeError"), "Revoked Proxy Apply TypeError");
}
