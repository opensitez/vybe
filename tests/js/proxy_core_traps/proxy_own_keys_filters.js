// vybe-test: js/proxy_core_traps/proxy_own_keys_filters
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

const handler = {
    ownKeys(target) {
        return Object.keys(target).filter(k => !k.startsWith("_"));
    }
};
const obj = new Proxy({ a: 1, _private: 2, b: 3 }, handler);
// Reflect.ownKeys must include all keys for proxy to work,
// but getOwnPropertyDescriptor will be called for each key returned
// For simplicity test that the filtered ownKeys works for Object.keys
// by also providing getOwnPropertyDescriptor
const handler2 = {
    ...handler,
    getOwnPropertyDescriptor(target, prop) {
        return Object.getOwnPropertyDescriptor(target, prop);
    }
};
const obj2 = new Proxy({ a: 1, _private: 2, b: 3 }, handler2);
const keys = Object.keys(obj2);
console.log(keys.join(","));
