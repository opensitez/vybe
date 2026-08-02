// vybe-test: js/proxy_traps_advanced/proxy_set_trap_validates_input
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

const proxy = new Proxy({}, {
    set(target, prop, value) {
        if (typeof value !== "number") throw new TypeError("must be number");
        target[prop] = value;
        return true;
    }
});
proxy.x = 42;
__check(__line(proxy.x), "42");
let threw = false;
try { proxy.y = "string"; } catch (e) { threw = e instanceof TypeError; }
__check(__line(threw), "true");
