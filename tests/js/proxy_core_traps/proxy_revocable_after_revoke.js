// vybe-test: js/proxy_core_traps/proxy_revocable_after_revoke
// origin: languages/js/tests/js/test_proxy_core_traps.rs

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
try { proxy.x; } catch { threw = true; }
__check(__line(threw), "true");
