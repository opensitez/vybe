// vybe-test: js/proxy_traps_advanced/proxy_set_trap_can_normalize_values
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
        target[prop] = typeof value === "string" ? value.toLowerCase() : value;
        return true;
    }
});
proxy.name = "HELLO";
__check(__line(proxy.name), "hello");
