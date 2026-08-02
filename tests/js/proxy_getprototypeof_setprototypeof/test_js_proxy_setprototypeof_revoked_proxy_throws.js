// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_setprototypeof_revoked_proxy_throws
// origin: languages/js/tests/js/test_js_proxy_getprototypeof_setprototypeof.rs

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

const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.setPrototypeOf(proxy, {});
} catch (e) {
    __check(__line("Revoked Proxy setPrototypeOf TypeError"), "Revoked Proxy setPrototypeOf TypeError");
}
