// vybe-test: js/proxy_traps_advanced/proxy_revocable_access_after_revoke_throws
// origin: languages/js/tests/js/test_proxy_traps_advanced.rs

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

const { proxy, revoke } = Proxy.revocable({ x: 1 }, {});
__check(__line(proxy.x), "1");
revoke();
let threw = false;
try { proxy.x; } catch (e) { threw = e instanceof TypeError; }
__check(__line(threw), "true");
