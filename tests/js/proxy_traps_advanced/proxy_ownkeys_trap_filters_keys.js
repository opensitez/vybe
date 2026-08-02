// vybe-test: js/proxy_traps_advanced/proxy_ownkeys_trap_filters_keys
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

const target = { a: 1, _private: 2, b: 3 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return Object.keys(t).filter(k => !k.startsWith("_"));
    }
});
__check(__line(Object.keys(proxy).sort().join(",")), "a,b");
