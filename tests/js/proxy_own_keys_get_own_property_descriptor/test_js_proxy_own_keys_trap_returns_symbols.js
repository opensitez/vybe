// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_trap_returns_symbols
// origin: languages/js/tests/js/test_js_proxy_own_keys_get_own_property_descriptor.rs

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

const s1 = Symbol("s1");
const s2 = Symbol("s2");
const target = { [s1]: 10, [s2]: 20 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return [s1]; // Filter out s2
    }
});
__check(__line(Object.getOwnPropertySymbols(proxy).length), "1");
