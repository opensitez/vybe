// vybe-test: js/proxy_traps_advanced/proxy_ownkeys_intercepts_object_getownpropertynames
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

const proxy = new Proxy({ x: 1, y: 2 }, {
    ownKeys() { return ["x", "z"]; }
});
// ownKeys must return subset of actual keys (invariant)
// or keys that exist in target for non-configurable
const keys = Object.getOwnPropertyNames(proxy);
console.log(keys.includes("x"));
